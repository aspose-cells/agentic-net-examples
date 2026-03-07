using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the source XLSX workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and configure to exclude unused styles
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.ExcludeUnusedStyles = true; // default is true, set explicitly for clarity

        // Save the workbook as an HTML file using the configured options
        workbook.Save("output.html", htmlOptions);
    }
}