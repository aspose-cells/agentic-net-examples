using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Custom implementation of IFilePathProvider.
    // It converts the worksheet name to a lowercase file name with .html extension.
    internal class LowerCaseFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Ensure the sheet name is safe for a file name.
            string safeName = CellsHelper.CreateSafeSheetName(sheetName);
            // Convert to lowercase and append .html.
            return $"{safeName.ToLowerInvariant()}.html";
        }
    }

    public class IFilePathProviderDemo
    {
        public static void Run()
        {
            // Create a new workbook and add some worksheets with mixed‑case names.
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Summary";
            workbook.Worksheets.Add("DataSheet");
            workbook.Worksheets.Add("Report_2023");

            // Populate some data (optional, just for illustration).
            workbook.Worksheets["Summary"].Cells["A1"].PutValue("This is the summary sheet.");
            workbook.Worksheets["DataSheet"].Cells["A1"].PutValue("Data goes here.");
            workbook.Worksheets["Report_2023"].Cells["A1"].PutValue("Report content.");

            // Configure HTML save options to use the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export each worksheet to a separate HTML file.
                ExportActiveWorksheetOnly = false,
                // Use the custom provider so that links reference lowercase file names.
                FilePathProvider = new LowerCaseFilePathProvider()
            };

            // Save the workbook. The main HTML file will reference
            // summary.html, datasheet.html, and report_2023.html.
            workbook.Save("WorkbookWithLowerCaseLinks.html", saveOptions);

            Console.WriteLine("Workbook saved with custom lowercase file paths.");
        }
    }

    // Entry point for testing.
    class Program
    {
        static void Main()
        {
            IFilePathProviderDemo.Run();
        }
    }
}