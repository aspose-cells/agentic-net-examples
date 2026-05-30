using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["B2"].PutValue(12345);

            // Define the output folder (relative to the current directory)
            string outputFolder = Path.Combine(Environment.CurrentDirectory, "HtmlOutput");
            // Define a subfolder for the generated CSS files
            string cssFolder = Path.Combine(outputFolder, "css");

            // Ensure both folders exist; CreateDirectory will create them if they do not exist
            Directory.CreateDirectory(cssFolder);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export each worksheet's CSS to a separate file
                ExportWorksheetCSSSeparately = true,
                // Specify the folder where the CSS files (and other attached resources) will be placed
                AttachedFilesDirectory = cssFolder,
                // Automatically create the output directory if it does not exist
                CreateDirectory = true
            };

            // Save the workbook as HTML; the CSS files will be written to the specified cssFolder
            string htmlPath = Path.Combine(outputFolder, "Workbook.html");
            workbook.Save(htmlPath, saveOptions);

            Console.WriteLine($"HTML file saved to: {htmlPath}");
            Console.WriteLine($"CSS files saved to: {cssFolder}");
        }
    }
}