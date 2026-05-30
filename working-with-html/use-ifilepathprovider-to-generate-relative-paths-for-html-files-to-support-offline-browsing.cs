using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider that generates relative file names.
    // This enables offline browsing because the generated HTML files reference each other
    // using simple relative paths (e.g., "Sheet1.html").
    public class RelativePathProvider : IFilePathProvider
    {
        public string GetFullName(string sheetName)
        {
            // Return a relative file name based on the worksheet name.
            // No directory information is added, so the files will be placed
            // in the same folder as the main HTML file.
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and add sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue("Hello from Sheet1");

            // Add a second worksheet to demonstrate multiple files.
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["A1"].PutValue("Hello from Sheet2");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();

            // Use the custom relative path provider.
            saveOptions.FilePathProvider = new RelativePathProvider();

            // Ensure links are relative (default is false, set explicitly for clarity).
            saveOptions.IsFullPathLink = false;

            // Optional: specify a directory for attached resources (images, CSS, etc.).
            // Here we keep them in a subfolder named "resources".
            string resourcesDir = Path.Combine(Directory.GetCurrentDirectory(), "resources");
            Directory.CreateDirectory(resourcesDir);
            saveOptions.AttachedFilesDirectory = resourcesDir;

            // Save the workbook. The main file will be "Workbook.html" and each worksheet
            // will be saved as a separate HTML file using the relative names provided.
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "Workbook.html");
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook saved to {outputPath}");
            Console.WriteLine("Individual worksheet HTML files are generated with relative links for offline use.");
        }
    }
}