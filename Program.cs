using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace CertificateGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string csvPath = Path.Combine(baseDir, "Data.csv");
            string templatePath = Path.Combine(baseDir, "Template.docx");
            string outputFolder = Path.Combine(baseDir, "Output");

            Logger.Init(baseDir);
            Logger.Log("System initialized.");

            if (!File.Exists(csvPath) || !File.Exists(templatePath))
            {
                Logger.Log("ERROR: Missing Data.csv or Template.docx.");
                Console.ReadLine();
                return;
            }
            Directory.CreateDirectory(outputFolder);

            Logger.Log("Loading data...");

            var rawLines = File.ReadAllLines(csvPath, Encoding.UTF8);
            var qualifiedEmployees = rawLines
                .Skip(1)
                .Select(line => line.Split(','))
                .Where(cols => cols.Length >= 5 && !string.IsNullOrWhiteSpace(cols[0])
                               && int.TryParse(cols[3], out _) && int.TryParse(cols[4], out _))
                .Select(cols => new Employee(cols))
                .GroupBy(emp => emp.FullName).Select(g => g.First())
                .Where(emp => emp.FinalScore >= 70)
                .ToList();

            Logger.Log($"Found {qualifiedEmployees.Count} qualified employees. Starting Word...");

            Word.Application wordApp = new Word.Application { Visible = false };

            try
            {
                int successCount = 0;
                int failCount = 0;

                foreach (var emp in qualifiedEmployees)
                {
                    try
                    {
                        Logger.Log($"Processing: {emp.FullName}...");

                        string bodyText = emp.FinalScore > 90
                            ? $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {emp.FinalScore:F1}.\n" +
                              "נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית"
                            : "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";

                        var replacements = new Dictionary<string, string>
                        {
                            { "FullName", emp.FullName },
                            { "Department", emp.Department },
                            { "Phone", "050-0000000" },
                            { "Email", $"{emp.FirstName}@gmail.com" },
                            { "BodyText", bodyText }
                        };

                        string safeName = PdfGenerator.CleanFileName($"{emp.FirstName}_{emp.LastName}");
                        string outputPath = Path.Combine(outputFolder, $"{safeName}.pdf");

                        PdfGenerator.Generate(wordApp, templatePath, outputPath, replacements);

                        Logger.Log($"SUCCESS: {emp.FullName}");
                        successCount++;
                    }
                    catch (Exception innerEx)
                    {
                        Logger.Log($"ERROR processing {emp.FullName}: {innerEx.Message}");
                        failCount++;
                    }
                }
                Logger.Log($"Job Done. Success: {successCount}, Failed: {failCount}");
            }
            catch (Exception ex)
            {
                Logger.Log($"FATAL ERROR: {ex.Message}");
            }
            finally
            {
                if (wordApp != null) { wordApp.Quit(); System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp); }
                Logger.Log("Word application closed.");
            }

            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }
    }
}