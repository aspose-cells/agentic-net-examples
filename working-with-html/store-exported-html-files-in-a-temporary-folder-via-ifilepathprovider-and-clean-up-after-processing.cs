// Title: Export Excel to HTML with Aspose.Cells using a custom IFilePathProvider and automatic temp‑folder cleanup (C#)
// Description: Shows how to implement IFilePathProvider to write each worksheet as a separate HTML file into a unique temporary directory, assign it to HtmlSaveOptions, save the workbook, and delete the folder after processing.
// Keywords: Aspose.Cells | C# | HTML export | IFilePathProvider | temporary directory | temp folder cleanup | per‑sheet HTML | HtmlSaveOptions | Workbook.Save
// Common Searches: Aspose.Cells IFilePathProvider example | C# export Excel to HTML per sheet | save Aspose.Cells HTML to temporary folder | delete temporary files after Aspose.Cells export | how to clean up temp directory Aspose.Cells
// Developer Intent: Create a unique temporary directory for per‑worksheet HTML files during an Aspose.Cells export and ensure the directory is removed after use.
// Use Cases: Web API endpoint that zips per‑sheet HTML files for download, then deletes the temp folder. | Background service generating HTML reports, storing them temporarily before emailing, and cleaning up afterward. | Automated test suite that exports workbooks to HTML and needs a clean environment for each run. | Serverless function that must not leave residual files on the host after completing the export.
// AI Prompts: Write C# code that uses Aspose.Cells to export a workbook to HTML with each sheet saved in a temporary folder and the folder automatically deleted after export. | Explain how to modify the example to stream generated HTML files to a MemoryStream instead of writing them to disk. | Show robust error handling for folder creation and deletion in the Aspose.Cells HTML export scenario. | Provide a PowerShell script to locate and clean leftover Aspose.Cells temporary folders on a server.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExportTemp
{
    // Custom IFilePathProvider that stores each worksheet HTML file in a temporary folder
    // Shows how to implement IFilePathProvider to write each worksheet as a separate HTML file into a unique temporary directory, assign it to HtmlSaveOptions, save the workbook, and delete the folder after processing.
    internal class TempFilePathProvider : IFilePathProvider
    {
        private readonly string _tempFolder;

        public TempFilePathProvider()
        {
            // Create a unique temporary directory for this export session
            _tempFolder = Path.Combine(Path.GetTempPath(), "AsposeHtmlExport_" + Guid.NewGuid().ToString("N"));
            Directory.CreateDirectory(_tempFolder);
        }

        // Returns the full path for a given worksheet name
        public string GetFullName(string sheetName)
        {
            // Ensure the temporary folder exists
            if (!Directory.Exists(_tempFolder))
                Directory.CreateDirectory(_tempFolder);

            // Build the file name (e.g., Sheet1.html) inside the temp folder
            return Path.Combine(_tempFolder, $"{sheetName}.html");
        }

        // Expose the temporary folder so it can be cleaned up after export
        public string TempFolder => _tempFolder;
    }

    public static class HtmlExportWithTempFolder
    {
        public static void Run()
        {
            // Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");

            // Initialize HtmlSaveOptions and assign the custom TempFilePathProvider
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            TempFilePathProvider provider = new TempFilePathProvider();
            saveOptions.FilePathProvider = provider;

            // Save the workbook to HTML. The main HTML file will reference the per‑sheet files
            // generated in the temporary folder.
            string mainHtmlPath = Path.Combine(provider.TempFolder, "main.html");
            workbook.Save(mainHtmlPath, saveOptions);

            Console.WriteLine($"Main HTML saved to: {mainHtmlPath}");
            Console.WriteLine($"Per‑sheet files are located in: {provider.TempFolder}");

            // OPTIONAL: Process the generated files here (e.g., read, zip, send over network)

            // Clean up: delete the temporary folder and all its contents
            try
            {
                Directory.Delete(provider.TempFolder, true);
                Console.WriteLine("Temporary folder cleaned up successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error cleaning up temporary folder: {ex.Message}");
            }
        }
    }

    // Entry point for demonstration
    class Program
    {
        static void Main()
        {
            HtmlExportWithTempFolder.Run();
        }
    }
}
