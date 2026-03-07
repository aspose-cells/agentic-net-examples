using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Initialize HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Specify the default font to be used when the original font is unavailable
        saveOptions.DefaultFontName = "Arial";

        // Save the workbook as an HTML file with the specified options
        workbook.Save("output.html", saveOptions);
    }
}