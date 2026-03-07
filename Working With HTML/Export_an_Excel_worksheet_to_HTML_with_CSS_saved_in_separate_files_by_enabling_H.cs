using System;
using System.IO;
using Aspose.Cells;

class ExportExcelToHtmlWithSeparateCss
{
    static void Main()
    {
        // Path to the source Excel file (XLSX)
        string inputPath = "input.xlsx"; // TODO: set the correct path to your XLSX file

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and enable exporting worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Define the output HTML file path
        string outputPath = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputPath, saveOptions);

        Console.WriteLine($"HTML file with separate CSS saved to: {outputPath}");
    }
}