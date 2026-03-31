using System;
using Aspose.Cells;

class XlsxToHtmlConverter
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path where the HTML output will be saved
        string outputPath = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();

        // Set PresentationPreference to true for a more beautiful (BestFit) layout
        htmlOptions.PresentationPreference = true;

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine($"Workbook successfully converted to HTML: {outputPath}");
    }
}