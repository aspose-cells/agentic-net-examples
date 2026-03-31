using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Set the cross‑cell string handling.
        // The HtmlCrossType enum does not contain a Strikethrough value;
        // the closest valid option that changes rendering behavior is Cross.
        htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}