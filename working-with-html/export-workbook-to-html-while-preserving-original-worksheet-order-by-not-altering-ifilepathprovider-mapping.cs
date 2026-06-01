using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom file path provider that returns a simple file name based on the worksheet name.
    // This preserves the original worksheet order because the provider does not modify the mapping.
    public class PreserveOrderFilePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Return a file name that keeps the sheet name unchanged.
            // Example: "Sheet1.html", "Sheet2.html", etc.
            return $"{sheetName}.html";
        }
    }

    public class ExportWorkbookToHtml
    {
        public static void Run()
        {
            // Create a new workbook and add sample data.
            Workbook workbook = new Workbook();

            // First worksheet (default)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Data in Sheet1");

            // Add a second worksheet.
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Data in Sheet2");

            // Add a third worksheet.
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");
            sheet3.Cells["A1"].PutValue("Data in Sheet3");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Use the custom provider to keep the original worksheet order.
                FilePathProvider = new PreserveOrderFilePathProvider(),

                // Ensure the whole workbook is exported (not only the active sheet).
                ExportActiveWorksheetOnly = false
            };

            // Save the workbook to HTML. Each worksheet will be saved as a separate HTML file
            // with names matching the worksheet names, preserving the original order.
            workbook.Save("WorkbookOutput.html", saveOptions);
        }
    }

    // Entry point for demonstration.
    class Program
    {
        static void Main()
        {
            ExportWorkbookToHtml.Run();
            Console.WriteLine("Workbook exported to HTML while preserving worksheet order.");
        }
    }
}