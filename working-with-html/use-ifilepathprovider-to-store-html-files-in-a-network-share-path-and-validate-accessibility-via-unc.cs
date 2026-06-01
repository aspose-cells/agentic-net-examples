using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkShareDemo
{
    // Custom implementation of IFilePathProvider that returns UNC paths for each worksheet.
    public class NetworkShareFilePathProvider : IFilePathProvider
    {
        private readonly string _uncBasePath; // e.g. \\Server\Share\Folder\

        public NetworkShareFilePathProvider(string uncBasePath)
        {
            // Ensure the base path ends with a backslash for proper Path.Combine behavior.
            if (!uncBasePath.EndsWith("\\"))
                uncBasePath += "\\";

            _uncBasePath = uncBasePath;
        }

        // Returns the full UNC file name for a given worksheet.
        public string GetFullName(string sheetName)
        {
            // Combine base UNC path with sheet name and .html extension.
            return Path.Combine(_uncBasePath, $"{sheetName}.html");
        }

        // Helper method to verify that the UNC directory is reachable.
        public bool IsUncPathAccessible()
        {
            // Directory.Exists works with UNC paths.
            return Directory.Exists(_uncBasePath);
        }
    }

    public class Program
    {
        public static void Main()
        {
            // Define the UNC network share where HTML files will be stored.
            string uncPath = @"\\MyServer\SharedFolder\ExcelHtmlExport";

            // Validate that the UNC share is reachable before proceeding.
            if (!Directory.Exists(uncPath))
            {
                Console.WriteLine($"UNC path not accessible: {uncPath}");
                return;
            }

            // Create a workbook and add sample data.
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "Report";
            ws.Cells["A1"].PutValue("Hello");
            ws.Cells["A2"].PutValue("World");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Use full path links so that the main HTML references the UNC files directly.
            saveOptions.IsFullPathLink = true;

            // Assign the custom file path provider.
            var provider = new NetworkShareFilePathProvider(uncPath);
            saveOptions.FilePathProvider = provider;

            // Save the workbook; the main HTML file can be saved locally.
            string localMainHtml = Path.Combine(Path.GetTempPath(), "WorkbookMain.html");
            workbook.Save(localMainHtml, saveOptions);
            Console.WriteLine($"Main HTML saved to: {localMainHtml}");

            // Verify that the worksheet HTML file was created on the UNC share.
            string expectedWorksheetFile = provider.GetFullName(ws.Name);
            if (File.Exists(expectedWorksheetFile))
            {
                Console.WriteLine($"Worksheet HTML successfully created at UNC location: {expectedWorksheetFile}");
            }
            else
            {
                Console.WriteLine($"Failed to create worksheet HTML at UNC location: {expectedWorksheetFile}");
            }
        }
    }
}