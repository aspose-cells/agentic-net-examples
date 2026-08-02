// Title: Save Excel Worksheets as Individual HTML Files to a UNC Share Using Aspose.Cells IFilePathProvider (C#)
// Description: Demonstrates how to implement a custom IFilePathProvider that generates UNC file names for each worksheet, validates the network share’s accessibility, configures HtmlSaveOptions with full‑path links, saves the workbook to the share, and verifies the creation of per‑worksheet HTML files.
// Keywords: Aspose.Cells | IFilePathProvider | UNC path | C# HTML export | HtmlSaveOptions | network share | per worksheet HTML | full path links | validate UNC access | GitHub example | .NET | Aspose.Cells for .NET
// Common Searches: Aspose.Cells IFilePathProvider UNC example | Save Excel as HTML to network share C# | Validate UNC folder before exporting Aspose.Cells | Export each worksheet to separate HTML files | Full path links HtmlSaveOptions Aspose | GitHub Aspose.Cells HTML export UNC
// Developer Intent: Export each worksheet to its own HTML file on a UNC network share while confirming the share is reachable before saving.
// Use Cases: Generate per‑worksheet HTML reports on a central file server for intranet or web consumption. | Automate nightly workbook exports to a shared UNC folder, creating one HTML file per sheet. | Prevent runtime errors by checking UNC accessibility and permissions prior to HTML export.
// AI Prompts: Provide a C# sample that extends IFilePathProvider to create timestamped UNC file names for Aspose.Cells HTML export. | Explain how to catch and handle permission or connectivity errors when using HtmlSaveOptions with a custom UNC FilePathProvider. | Show how to modify the code to also export worksheet images to the same UNC folder and reference them with full‑path links.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsNetworkHtmlExport
{
    // Custom implementation of IFilePathProvider that generates UNC paths for each worksheet.
    // Demonstrates how to implement a custom IFilePathProvider that generates UNC file names for each worksheet, validates the network share’s accessibility, configures HtmlSaveOptions with full‑path links, saves the workbook to the share, and verifies the creation of per‑worksheet HTML files.
    public class UncFilePathProvider : IFilePathProvider
    {
        private readonly string _baseUncPath;

        public UncFilePathProvider(string baseUncPath)
        {
            // Ensure the base path does not end with a directory separator.
            _baseUncPath = baseUncPath.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
        }

        // Returns the full UNC file name for a given worksheet.
        public string GetFullName(string sheetName)
        {
            // Example: \\server\share\folder\Sheet1.html
            return Path.Combine(_baseUncPath, $"{sheetName}.html");
        }
    }

    class Program
    {
        static void Main()
        {
            // UNC network share where HTML files will be stored.
            string networkUncPath = @"\\myserver\share\excel_html";

            // Validate that the UNC path is accessible.
            if (!Directory.Exists(networkUncPath))
            {
                Console.WriteLine($"Error: The network path '{networkUncPath}' is not accessible.");
                return;
            }

            // Create a new workbook and add sample data.
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Summary";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");

            // Add a second worksheet to demonstrate separate HTML files.
            Worksheet sheet2 = workbook.Worksheets.Add("Details");
            sheet2.Cells["A1"].PutValue("Detail data");

            // Configure HTML save options.
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Use full path links so that the generated HTML references the UNC files correctly.
            saveOptions.IsFullPathLink = true;
            // Assign the custom UNC file path provider.
            saveOptions.FilePathProvider = new UncFilePathProvider(networkUncPath);

            // Save the workbook. The main HTML file can be placed in the same UNC folder.
            string mainHtmlPath = Path.Combine(networkUncPath, "WorkbookIndex.html");
            workbook.Save(mainHtmlPath, saveOptions);

            // Verify that each worksheet HTML file was created.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                string expectedPath = Path.Combine(networkUncPath, $"{ws.Name}.html");
                bool exists = File.Exists(expectedPath);
                Console.WriteLine($"Worksheet '{ws.Name}' HTML file exists: {exists}");
            }

            Console.WriteLine("Export completed.");
        }
    }
}
