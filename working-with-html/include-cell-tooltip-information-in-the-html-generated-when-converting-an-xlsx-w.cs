using System;
using Aspose.Cells;

namespace AsposeCellsTooltipExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with your actual file path)
            string inputPath = "input.xlsx";
            Workbook workbook = new Workbook(inputPath);

            // Optional: make a cell's content longer than its column width to demonstrate tooltip
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This is a very long text that will not fit in the column width and should show a tooltip.");
            sheet.Cells.SetColumnWidth(0, 10); // Narrow column width for column A

            // Configure HTML save options to include tooltip text for truncated cell data
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.AddTooltipText = true; // Enable tooltip generation

            // Save the workbook as HTML with the tooltip option enabled
            string outputPath = "output_with_tooltip.html";
            workbook.Save(outputPath, htmlOptions);

            Console.WriteLine($"HTML file saved with tooltips at: {outputPath}");
        }
    }
}