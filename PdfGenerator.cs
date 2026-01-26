using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace CertificateGenerator
{
    public static class PdfGenerator
    {
        public static void Generate(Word.Application app, string templatePath, string outputPath, Dictionary<string, string> data)
        {
            Word.Document doc = null;
            try
            {
                doc = app.Documents.Open(templatePath, ReadOnly: true);

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
                doc.ExportAsFixedFormat(outputPath, Word.WdExportFormat.wdExportFormatPDF);
            }
            finally
            {
                if (doc != null) doc.Close(false);
            }
        }

        public static string CleanFileName(string fileName)
        {
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }
            return fileName;
        }
    }
}