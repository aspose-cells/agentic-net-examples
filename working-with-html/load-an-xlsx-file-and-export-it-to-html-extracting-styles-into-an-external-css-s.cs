using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the source XLSX file
        string sourcePath = "input.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(sourcePath);

        // Configure HTML save options to export the worksheet CSS separately
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true; // CSS will be written to an external .css file

        // Define the output HTML file path
        string outputHtml = "output.html";

        // Save the workbook as HTML using the configured options
        workbook.Save(outputHtml, saveOptions);

        // Inform the user where the files were saved
        Console.WriteLine("HTML file saved to: " + Path.GetFullPath(outputHtml));
        Console.WriteLine("Associated CSS file saved alongside the HTML file.");
    }
}