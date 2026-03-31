using System;
using Aspose.Cells;

class HtmlExportWithPresentationPreference
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Path for the resulting HTML file
        string outputPath = "output.html";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Create HTML save options and enable presentation preference (best‑fit layout)
        HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
        htmlOptions.PresentationPreference = true;

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, htmlOptions);

        Console.WriteLine("Workbook has been exported to HTML with PresentationPreference enabled.");
    }
}