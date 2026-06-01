using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    // Provides virtual file names for each worksheet when exporting to HTML.
    public class CustomFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Simple file name; in a real scenario this could be a URL.
            return $"{sheetName}.html";
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // Generate HTML for the whole workbook and save to a file.
                string workbookHtml = GetWorkbookHtml();
                File.WriteAllText("Workbook.html", workbookHtml);
                Console.WriteLine("Workbook HTML saved to Workbook.html");

                // Example: generate HTML for a single worksheet.
                string sheetHtml = GetWorksheetHtml("Summary");
                File.WriteAllText("Summary.html", sheetHtml);
                Console.WriteLine("Worksheet HTML saved to Summary.html");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors.
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        // Creates a sample workbook and returns its HTML representation.
        private static string GetWorkbookHtml()
        {
            // Create a new workbook with two worksheets.
            var workbook = new Workbook();
            var sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Summary";
            sheet1.Cells["A1"].PutValue("Hello Aspose.Cells!");
            sheet1.Cells["A2"].PutValue(DateTime.Now);

            var sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Detail data");
            sheet2.Cells["B2"].PutValue(12345);

            // Configure HTML export options.
            var saveOptions = new HtmlSaveOptions
            {
                FilePathProvider = new CustomFilePathProvider(),
                SaveAsSingleFile = true,
                ExportActiveWorksheetOnly = false
            };

            // Save to a memory stream and return the HTML string.
            using var ms = new MemoryStream();
            workbook.Save(ms, saveOptions);
            ms.Position = 0;
            using var reader = new StreamReader(ms);
            return reader.ReadToEnd();
        }

        // Exports only the specified worksheet to HTML.
        private static string GetWorksheetHtml(string sheetName)
        {
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = sheetName;
            sheet.Cells["A1"].PutValue($"Content of sheet '{sheetName}'");

            var saveOptions = new HtmlSaveOptions
            {
                FilePathProvider = new CustomFilePathProvider(),
                ExportActiveWorksheetOnly = true,
                SaveAsSingleFile = true
            };

            using var ms = new MemoryStream();
            workbook.Save(ms, saveOptions);
            ms.Position = 0;
            using var reader = new StreamReader(ms);
            return reader.ReadToEnd();
        }

        // Loads a workbook from a file if it exists; otherwise throws a clear exception.
        private static Workbook LoadWorkbookIfExists(string path)
        {
            if (!File.Exists(path))
                throw new FileNotFoundException($"Template file not found: {path}");
            return new Workbook(path);
        }
    }
}