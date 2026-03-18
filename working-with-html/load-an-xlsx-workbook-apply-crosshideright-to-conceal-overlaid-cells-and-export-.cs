using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and set the cross‑cell string handling to CrossHideRight
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.HtmlCrossStringType = HtmlCrossType.CrossHideRight; // hides the right part when strings overlap

        // Export the workbook to HTML using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}