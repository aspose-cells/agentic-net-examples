using System;
using Aspose.Cells;

class ExportWorksheetToHtml
{
    static void Main()
    {
        // Load the source XLSX workbook
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Configure HTML save options to export worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true;

        // Define the output HTML file path
        string outputHtml = "output.html";

        // Save the workbook as HTML; CSS will be written to separate files
        workbook.Save(outputHtml, saveOptions);

        Console.WriteLine("Workbook exported to HTML with separate CSS files.");
    }
}