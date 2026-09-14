// Title: Export an Aspose.Cells workbook to HTML on a UNC network share using a custom IFilePathProvider with accessibility validation
// AI Prompts: Implement a class that inherits IFilePathProvider to return a UNC path and assign it to HtmlSaveOptions.FilePathProvider for HTML export. | Add logic to verify that a UNC directory exists and is writable before invoking Workbook.Save with HtmlSaveOptions. | Extend the sample to create additional HTML assets (images, CSS files) on the same network share via the custom file path provider.
// Common Searches: how to save Aspose.Cells workbook as HTML to a UNC network share in C# | C# check write permission on UNC folder before exporting HTML with Aspose.Cells | custom IFilePathProvider example for Aspose.Cells HTML output | Aspose.Cells HtmlSaveOptions UNC path usage | error handling for inaccessible network share when saving HTML with Aspose.Cells
// Tags: IFilePathProvider UNC implementation | HtmlSaveOptions network share | Aspose.Cells HTML export to UNC | validate UNC folder write access C# | custom file path provider Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkShareExample
{
    // Custom file path provider that returns a UNC network share path for HTML files.
    // The example defines a NetworkShareFilePathProvider that implements IFilePathProvider to prepend a UNC root to HTML file names, checks the UNC directory's existence and write permissions, creates a sample workbook, and saves it as HTML to the network share using HtmlSaveOptions.
    public class NetworkShareFilePathProvider : IFilePathProvider
    {
        private readonly string _networkShareRoot;

        public NetworkShareFilePathProvider(string networkShareRoot)
        {
            // Ensure the root path ends with a backslash.
            _networkShareRoot = networkShareRoot.EndsWith("\\") ? networkShareRoot : networkShareRoot + "\\";
        }

        // Called by Aspose.Cells when it needs to resolve a file name.
        public string GetFilePath(string fileName)
        {
            // Combine the UNC root with the requested file name.
            return Path.Combine(_networkShareRoot, fileName);
        }

        // Required by IFilePathProvider – return the full path for the given file name.
        public string GetFullName(string fileName)
        {
            return GetFilePath(fileName);
        }
    }

    class Program
    {
        static void Main()
        {
            // UNC path to the network share where HTML files will be stored.
            string networkSharePath = @"\\MyServer\SharedFolder\AsposeHtmlOutput";

            // Validate that the UNC path is accessible.
            if (!Directory.Exists(networkSharePath))
            {
                Console.WriteLine($"Error: The network share path '{networkSharePath}' is not accessible.");
                return;
            }

            // Optional: test write permission by creating a temporary file.
            string testFilePath = Path.Combine(networkSharePath, "write_test.tmp");
            try
            {
                File.WriteAllText(testFilePath, "test");
                File.Delete(testFilePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: Write permission denied on '{networkSharePath}'. Details: {ex.Message}");
                return;
            }

            // Create a new workbook and add some sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(85);

            // Configure HTML save options to use the custom file path provider.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions(SaveFormat.Html);
            saveOptions.FilePathProvider = new NetworkShareFilePathProvider(networkSharePath);

            // Define the HTML file name (without path, provider will prepend the UNC root).
            string htmlFileName = "Report.html";

            // Save the workbook as HTML to the network share.
            try
            {
                workbook.Save(htmlFileName, saveOptions);
                Console.WriteLine($"Workbook successfully saved as HTML to UNC path: {Path.Combine(networkSharePath, htmlFileName)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
    }
}
