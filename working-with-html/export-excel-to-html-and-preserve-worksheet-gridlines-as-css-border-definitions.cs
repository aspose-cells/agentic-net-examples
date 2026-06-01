using System;
using Aspose.Cells;

namespace AsposeCellsExportGridlines
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Make gridlines visible in the worksheet (optional, but mirrors Excel UI)
            sheet.IsGridlinesVisible = true;

            // Populate some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.25);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.80);

            // Configure HTML save options to export gridlines as CSS borders
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExportGridLines = true,               // Preserve gridlines
                ExportActiveWorksheetOnly = true      // Export only the active sheet (optional)
            };

            // Save the workbook as HTML with the specified options
            string outputPath = "ExportedWithGridlines.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"Workbook exported to HTML with gridlines at: {outputPath}");
        }
    }
}