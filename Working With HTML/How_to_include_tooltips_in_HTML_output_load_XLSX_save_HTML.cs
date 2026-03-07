using System;
using Aspose.Cells;

namespace AsposeCellsTooltipExample
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook
            Workbook workbook = new Workbook("input.xlsx");

            // Optional: make a cell's content longer than its column width to see the tooltip effect
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("This is a very long text that will not fit in the column width and should show a tooltip");
            sheet.Cells.SetColumnWidth(0, 10); // narrow column

            // Configure HTML save options to add tooltip text when data is truncated
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
            htmlOptions.AddTooltipText = true; // enable tooltips

            // Save the workbook as HTML with tooltips enabled
            workbook.Save("output_with_tooltips.html", htmlOptions);

            Console.WriteLine("HTML file saved with tooltips.");
        }
    }
}