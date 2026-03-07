using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string sourceFile = "input.xlsx";

        // Path where the resulting HTML file will be saved
        string htmlFile = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourceFile);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Expand cell text beyond column width (ignores column width)
        // This makes overflowing text visible when exported to HTML,
        // effectively showing the text expanding from right to left.
        htmlOptions.FormatDataIgnoreColumnWidth = true;

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlFile, htmlOptions);
    }
}