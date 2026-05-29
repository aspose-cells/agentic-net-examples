using System;
using System.IO;
using Aspose.Cells;

class BatchHtmlExport
{
    static void Main()
    {
        // Folder containing the source XLSX files
        string inputFolder = @"C:\InputXlsx";

        // Folder where the reduced‑size HTML files will be saved
        string outputFolder = @"C:\OutputHtml";

        // Ensure the output directory exists
        Directory.CreateDirectory(outputFolder);

        // Retrieve all XLSX files in the input folder
        string[] xlsxFiles = Directory.GetFiles(inputFolder, "*.xlsx", SearchOption.TopDirectoryOnly);

        foreach (string xlsxPath in xlsxFiles)
        {
            // Load the workbook from the current XLSX file
            Workbook workbook = new Workbook(xlsxPath);

            // Create HTML save options and explicitly exclude unused styles
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions();
            htmlOptions.ExcludeUnusedStyles = true; // reduces HTML size

            // Build the output HTML file path (same name, .html extension)
            string htmlFileName = Path.GetFileNameWithoutExtension(xlsxPath) + ".html";
            string htmlPath = Path.Combine(outputFolder, htmlFileName);

            // Save the workbook as HTML using the configured options
            workbook.Save(htmlPath, htmlOptions);
        }

        Console.WriteLine("Batch conversion of XLSX files to HTML completed.");
    }
}