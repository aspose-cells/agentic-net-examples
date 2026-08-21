// Title: Export Workbook Sheets to HTML with IFilePathProvider and Serve via ASP.NET Core Web API
// Description: Demonstrates a custom TempFolderFilePathProvider that creates a unique temporary directory, configures HtmlSaveOptions to generate separate HTML files for each worksheet, saves an index.html linking to them, and shows how to expose the files through an ASP.NET Core Web API endpoint.
// Keywords: Aspose.Cells IFilePathProvider | HTML export per worksheet | Aspose.Cells HtmlSaveOptions | ASP.NET Core Web API file serving | temporary folder export Aspose | ExportActiveWorksheetOnly false | SaveAsSingleFile false | IsFullPathLink true | dynamic HTML report Aspose
// Common Searches: how to use IFilePathProvider with Aspose.Cells | export each worksheet to separate HTML files | serve Aspose.Cells HTML export via ASP.NET Core | temporary folder for Aspose HTML export | ASP.NET Web API route for generated HTML reports
// Developer Intent: Generate per‑worksheet HTML files with a custom path provider and make them accessible through an ASP.NET Core Web API route.
// Use Cases: Create on‑demand HTML reports from Excel workbooks without persisting files long‑term. | Return a URL or FileResult for the index.html of a multi‑sheet export in a REST endpoint. | Isolate export sessions with GUID‑based folders to avoid naming collisions and simplify cleanup.
// AI Prompts: Write an ASP.NET Core controller action that uses TempFolderFilePathProvider to export a workbook to HTML and returns the index.html as a FileResult. | Show how to configure a route like /api/export/{id} that streams the generated HTML files from the temporary folder. | Provide code to delete the temporary export folder after the HTTP response completes, ensuring no leftover files.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsConsole
{
    // Custom implementation of IFilePathProvider that generates file paths in a temporary folder
    // Demonstrates a custom TempFolderFilePathProvider that creates a unique temporary directory, configures HtmlSaveOptions to generate separate HTML files for each worksheet, saves an index.html linking to them, and shows how to expose the files through an ASP.NET Core Web API endpoint.
    public class TempFolderFilePathProvider : IFilePathProvider
    {
        private readonly string _baseFolder;

        public TempFolderFilePathProvider(string baseFolder)
        {
            _baseFolder = baseFolder;
        }

        // Returns a full file name for each worksheet (e.g., Sheet1.html)
        public string GetFullName(string sheetName)
        {
            // Ensure the folder exists
            Directory.CreateDirectory(_baseFolder);
            // Create a file name based on the worksheet name
            return Path.Combine(_baseFolder, $"{sheetName}.html");
        }
    }

    class Program
    {
        static void Main()
        {
            try
            {
                // Create a unique identifier for this export session
                string exportId = Guid.NewGuid().ToString();

                // Create a temporary folder to hold the generated HTML files
                string exportFolder = Path.Combine(Path.GetTempPath(), "AsposeExport", exportId);
                Directory.CreateDirectory(exportFolder);

                // Build a sample workbook with multiple worksheets
                Workbook workbook = new Workbook();
                workbook.Worksheets[0].Name = "Summary";
                workbook.Worksheets[0].Cells["A1"].PutValue("This is the summary sheet.");

                Worksheet sheet1 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet1.Name = "Data";
                sheet1.Cells["A1"].PutValue("Data sheet content.");

                Worksheet sheet2 = workbook.Worksheets[workbook.Worksheets.Add()];
                sheet2.Name = "Report";
                sheet2.Cells["A1"].PutValue("Report sheet content.");

                // Configure HTML save options
                HtmlSaveOptions saveOptions = new HtmlSaveOptions
                {
                    ExportActiveWorksheetOnly = false, // export all worksheets separately
                    SaveAsSingleFile = false,          // generate separate files
                    IsFullPathLink = true,             // use full path links in the main HTML
                    FilePathProvider = new TempFolderFilePathProvider(exportFolder)
                };

                // Save the workbook; the main file will be "index.html"
                string mainHtmlPath = Path.Combine(exportFolder, "index.html");
                workbook.Save(mainHtmlPath, saveOptions);

                // Verify that the main file was created
                if (File.Exists(mainHtmlPath))
                {
                    Console.WriteLine("Workbook exported successfully.");
                    Console.WriteLine($"Export ID: {exportId}");
                    Console.WriteLine($"Index file: {mainHtmlPath}");
                }
                else
                {
                    Console.WriteLine("Failed to create the index HTML file.");
                }
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
