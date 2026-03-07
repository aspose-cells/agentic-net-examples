using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Create HTML save options and enable exclusion of unused styles
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExcludeUnusedStyles = true; // This reduces the HTML file size

        // Save the workbook as an HTML file using the specified options
        workbook.Save("output.html", saveOptions);
    }
}