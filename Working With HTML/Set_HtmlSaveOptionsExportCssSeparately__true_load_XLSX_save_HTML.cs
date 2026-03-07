using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();

        // Export the worksheet CSS to a separate file
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Path for the generated HTML file
        string outputPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);
    }
}