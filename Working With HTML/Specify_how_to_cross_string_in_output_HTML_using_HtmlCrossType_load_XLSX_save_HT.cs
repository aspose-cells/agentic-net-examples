using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set the cross‑cell string handling
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        // Options: Default, MSExport, Cross, CrossHideRight, FitToCell
        // Here we use Cross for fast rendering of large files
        htmlOptions.HtmlCrossStringType = HtmlCrossType.Cross;

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}