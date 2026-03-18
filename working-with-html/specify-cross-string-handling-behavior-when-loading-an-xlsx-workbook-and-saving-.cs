using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load an existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and specify how cross‑cell strings are handled
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Example: use CrossHideRight to display cross‑cell strings and hide the right‑hand overlap.
        // Other options include Default, MSExport, Cross, FitToCell.
        htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}