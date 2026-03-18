using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCustomHtmlExport
{
    // Custom implementation of IFilePathProvider to control the file name of each worksheet HTML file
    public class CustomFilePathProvider : IFilePathProvider
    {
        // Returns a custom file name based on the worksheet name
        public string GetFullName(string sheetName)
        {
            // Example: place each sheet HTML file in a subfolder named "Sheets"
            // and prefix the file name with "custom_"
            string folder = Path.Combine("output", "html", "Sheets");
            Directory.CreateDirectory(folder); // Ensure the folder exists
            return Path.Combine(folder, $"custom_{sheetName}.html");
        }
    }

    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook (can be any custom location)
            string sourcePath = Path.Combine("input", "SampleWorkbook.xlsx");

            // Ensure the source workbook exists; create a simple one if it does not
            if (!File.Exists(sourcePath))
            {
                Directory.CreateDirectory(Path.GetDirectoryName(sourcePath));
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Name = "Sheet1";
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample Data");
                tempWb.Save(sourcePath);
            }

            // Load the workbook from the specified location
            Workbook workbook = new Workbook(sourcePath);

            // Configure HTML save options with custom output locations
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions
            {
                // Directory where attached files (images, etc.) will be saved
                AttachedFilesDirectory = Path.Combine("output", "html", "AttachedFiles"),

                // URL prefix used in the generated HTML to reference attached files
                AttachedFilesUrlPrefix = "AttachedFiles/",

                // Ensure that the directories are created automatically if they do not exist
                CreateDirectory = true,

                // Use the custom file path provider to control per‑sheet HTML file names
                FilePathProvider = new CustomFilePathProvider()
            };

            // Ensure the main HTML file is saved to a custom location
            string mainHtmlPath = Path.Combine("output", "html", "Workbook.html");
            Directory.CreateDirectory(Path.GetDirectoryName(mainHtmlPath));

            // Save the workbook as HTML using the configured options
            workbook.Save(mainHtmlPath, htmlOptions);

            Console.WriteLine("Workbook exported to HTML with custom output locations.");
        }
    }
}