using System;
using System.IO;
using Aspose.Cells;

class SeparateCssDemo
{
    static void Main()
    {
        // Path to the source XLSX file
        string inputPath = "input.xlsx";

        // Load the workbook from the XLSX file
        Workbook workbook = new Workbook(inputPath);

        // Create HTML save options and enable separate CSS export
        HtmlSaveOptions saveOptions = new HtmlSaveOptions();
        saveOptions.ExportWorksheetCSSSeparately = true; // CSS will be written to separate files

        // Define a folder where the HTML file and its related CSS (and images) will be stored
        string outputFolder = "HtmlOutput";
        Directory.CreateDirectory(outputFolder);
        saveOptions.AttachedFilesDirectory = outputFolder; // Folder for CSS and other attached files

        // Path for the main HTML file
        string htmlPath = Path.Combine(outputFolder, "output.html");

        // Save the workbook as HTML using the configured options
        workbook.Save(htmlPath, saveOptions);

        Console.WriteLine($"HTML file saved to: {htmlPath}");
        Console.WriteLine($"Separate CSS files are located in: {outputFolder}");
    }
}