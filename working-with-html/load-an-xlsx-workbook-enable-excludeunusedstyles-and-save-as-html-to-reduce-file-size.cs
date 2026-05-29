using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX workbook
        string inputPath = "input.xlsx";

        // Path where the HTML file will be saved
        string outputPath = "output.html";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and enable exclusion of unused styles
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExcludeUnusedStyles = true; // Reduces HTML file size

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);
    }
}