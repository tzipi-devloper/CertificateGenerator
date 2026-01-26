using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CertificateGenerator
{
    class Program
    {
        static string logPath;

        static void Main(string[] args)
        {
            string baseDir = AppDomain.CurrentDomain.BaseDirectory;
            string csvPath = Path.Combine(baseDir, "Data.csv");
            string templatePath = Path.Combine(baseDir, "Template.docx");
            string outputFolder = Path.Combine(baseDir, "Output");

            logPath = Path.Combine(baseDir, "application.log");
            File.WriteAllText(logPath, $"--- Process Started: {DateTime.Now} ---\n");

            Log("System initialized. Checking files...");

            if (!File.Exists(csvPath) || !File.Exists(templatePath))
            {
                Log("ERROR: Missing Data.csv or Template.docx.");
                Console.ReadLine();
                return;
            }
            Directory.CreateDirectory(outputFolder);

            Log("Loading data from CSV...");

            var qualifiedEmployees = File.ReadAllLines(csvPath)
                .Skip(1)
                .Select(line => line.Split(','))
                .Where(cols => cols.Length >= 5 && !string.IsNullOrWhiteSpace(cols[0]))
                .Select(cols => new Employee(cols))
                .GroupBy(emp => emp.FullName)
                .Select(g => g.First())
                .Where(emp => emp.FinalScore >= 70)
                .ToList();

            Log($"Found {qualifiedEmployees.Count} qualified employees. Starting Word...");

            Word.Application wordApp = new Word.Application { Visible = false };

            try
            {
                int successCount = 0;
                int failCount = 0;

                foreach (var emp in qualifiedEmployees)
                {
                    // === התיקון החשוב: הגנה מפני קריסה (Resilience) ===
                    try
                    {
                        Log($"Processing: {emp.FullName} (Score: {emp.FinalScore:F1})...");

                        string bodyText = emp.FinalScore > 90
                            ? $"הרינו להודיעך כי עברת בהצלחה את ההכשרה. הציון הסופי שלך הינו {emp.FinalScore:F1}.\n" +
                              "נמצאת מתאימ/ה לתפקיד מוביל/ה טכנולוגי מחלקתית"
                            : "הרינו להודיעך כי לא עברת את ההכשרה אך לצערנו לא נמצא תפקיד מתאים עבורך.";

                        var replacements = new Dictionary<string, string>
                        {
                            { "FullName", emp.FullName },
                            { "Department", emp.Department },
                            { "Phone", "050-0000000" },
                            { "Email", $"{emp.FirstName}@hogery.com" },
                            { "BodyText", bodyText }
                        };

                        GeneratePdf(wordApp, templatePath, Path.Combine(outputFolder, $"{emp.FirstName}_{emp.LastName}.pdf"), replacements);
                        Log($"SUCCESS: {emp.FullName}");
                        successCount++;
                    }
                    catch (Exception innerEx)
                    {
                        // אם יש תקלה בעובד אחד - רושמים לוג וממשיכים לאחרים!
                        Log($"ERROR processing {emp.FullName}: {innerEx.Message}");
                        failCount++;
                    }
                }

                Log($"Job Done. Success: {successCount}, Failed: {failCount}");
            }
            catch (Exception ex)
            {
                Log($"FATAL ERROR: {ex.Message}");
            }
            finally
            {
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                Log("Word application closed.");
            }

            Console.WriteLine("Press Enter to exit.");
            Console.ReadLine();
        }

        static void Log(string message)
        {
            string entry = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} | {message}";
            Console.WriteLine(entry);
            try { File.AppendAllText(logPath, entry + Environment.NewLine); } catch { }
        }

        static void GeneratePdf(Word.Application app, string template, string output, Dictionary<string, string> data)
        {
            Word.Document doc = app.Documents.Open(template, ReadOnly: true);
            foreach (Word.Field field in doc.Fields)
            {
                if (field.Code.Text.Contains("MERGEFIELD"))
                {
                    string fieldName = field.Code.Text.Split(new[] { "MERGEFIELD" }, StringSplitOptions.None)[1].Trim().Split(' ')[0].Trim('"');
                    if (data.TryGetValue(fieldName, out string value)) { field.Select(); app.Selection.TypeText(value); }
                }
            }
            doc.ExportAsFixedFormat(output, Word.WdExportFormat.wdExportFormatPDF);
            doc.Close(false);
        }
    }
}