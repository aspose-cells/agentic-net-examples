using System;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main(string[] args)
    {
        // Folder containing the XLSX files; can be passed as a command‑line argument
        string inputFolder = args.Length > 0 ? args[0] : @"C:\InputFolder";

        // Folder where the generated HTML files will be placed
        string outputFolder = Path.Combine(inputFolder, "ProcessedHtml");
        Directory.CreateDirectory(outputFolder);

        // Process each .xlsx file in the folder
        foreach (string xlsxPath in Directory.GetFiles(inputFolder, "*.xlsx"))
        {
            // Determine if the file size exceeds 5 MB
            bool excludeUnusedStyles = new FileInfo(xlsxPath).Length > 5L * 1024 * 1024;

            // Load the workbook (uses the provided Workbook(string) constructor)
            Workbook workbook = new Workbook(xlsxPath);

            // Configure HTML save options; set ExcludeUnusedStyles only for large files
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                ExcludeUnusedStyles = excludeUnusedStyles
            };

            // Build the output HTML file name
            string htmlPath = Path.Combine(
                outputFolder,
                Path.GetFileNameWithoutExtension(xlsxPath) + ".html");

            // Save the workbook as HTML using the provided Save(string, SaveOptions) method
            workbook.Save(htmlPath, htmlOptions);
        }
    }
}