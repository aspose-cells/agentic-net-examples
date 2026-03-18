using System;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Path to the source XLSX workbook
        string inputPath = "input.xlsx";

        // Path where the resulting HTML file will be saved
        string outputPath = "output.html";

        // Load the existing workbook from the file system
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and enable tooltip text preservation
        HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
        saveOptions.AddTooltipText = true; // Preserve tooltips for cells with overflowed data

        // Save the workbook as an HTML file using the configured options
        workbook.Save(outputPath, saveOptions);
    }
}