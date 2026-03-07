using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable simulation of unsupported borders
        // (ExportSimilarBorderStyle corresponds to SimulateUnsupportedBorders)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html);
        htmlOptions.ExportSimilarBorderStyle = true;

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", htmlOptions);
    }
}