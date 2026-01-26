using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace CertificateGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            string folderPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Work");
            string csvPath = Path.Combine(folderPath, "Data.csv");
            string templatePath = Path.Combine(folderPath, "Template.docx");
            string outputFolder = Path.Combine(folderPath, "Output");

            if (!File.Exists(csvPath) || !File.Exists(templatePath))
            {
                Console.WriteLine("Error: Missing files in 'Work' folder.");
                Console.ReadLine();
                return;
            }
            Directory.CreateDirectory(outputFolder);

            Console.WriteLine("Loading and processing data using LINQ...");

       
            var qualifiedEmployees = File.ReadAllLines(csvPath)
                .Skip(1) 
                .Select(line => line.Split(',')) 
                .Where(cols => cols.Length >= 5 && !string.IsNullOrWhiteSpace(cols[0]))
                .Select(cols => new Employee(cols)) 
                .GroupBy(emp => emp.FullName)
                .Select(g => g.First())
                .Where(emp => emp.FinalScore >= 70)
                .ToList();

            Console.WriteLine($"Found {qualifiedEmployees.Count} qualified employees. Generating PDFs...");

            Word.Application wordApp = new Word.Application { Visible = false };

            try
            {
                foreach (var emp in qualifiedEmployees)
                {
                    Console.WriteLine($"Generating for: {emp.FullName} (Score: {emp.FinalScore:F1})");


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

                    GeneratePdf(wordApp, templatePath, Path.Combine(outputFolder, $"{emp.FirstName}_{emp.LastName}.pdf"), replacements);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Critical Error: {ex.Message}");
            }
            finally
            {
                wordApp.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
            }

            Console.WriteLine("Process Complete. Press Enter to exit.");
            Console.ReadLine();
        }

      
        static void GeneratePdf(Word.Application app, string template, string output, Dictionary<string, string> data)
        {
            Word.Document doc = app.Documents.Open(template, ReadOnly: true);

            foreach (Word.Field field in doc.Fields)
            {
                if (field.Code.Text.Contains("MERGEFIELD"))
                {
                    string fieldName = field.Code.Text.Split(new[] { "MERGEFIELD" }, StringSplitOptions.None)[1].Trim().Split(' ')[0].Trim('"');

                    if (data.TryGetValue(fieldName, out string value))
                    {
                        field.Select();
                        app.Selection.TypeText(value);
                    }
                }
            }

            doc.ExportAsFixedFormat(output, Word.WdExportFormat.wdExportFormatPDF);
            doc.Close(false);

        }
    }
}