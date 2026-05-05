using System;
using Aspose.Cells;

namespace AsposeCellsConditionalFormattingHtml
{
    class Program
    {
        static void Main()
        {
            // Load the source XLSX workbook (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Create HTML save options
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

            // Preserve all conditional formatting styles (DataBar, ColorScale, IconSet)
            htmlOptions.ExcludeUnusedStyles = false;

            // Export worksheet CSS separately to ensure conditional formatting CSS is retained.
            htmlOptions.ExportWorksheetCSSSeparately = true;

            // Save the workbook as HTML with the configured options
            workbook.Save("output.html", htmlOptions);

            Console.WriteLine("Workbook converted to HTML with DataBar, ColorScale, and IconSet preserved.");
        }
    }
}