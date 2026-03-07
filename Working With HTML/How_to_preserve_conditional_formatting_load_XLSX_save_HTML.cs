using System;
using Aspose.Cells;

class PreserveConditionalFormatting
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Configure HTML save options to keep conditional formatting intact
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Merge conditional‑formatting areas before saving (preserves them in the output)
        htmlOptions.MergeAreas = true;

        // Do not exclude unused styles – this ensures all formatting, including conditional,
        // is written to the HTML file
        htmlOptions.ExcludeUnusedStyles = false;

        // Save the workbook as HTML with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}