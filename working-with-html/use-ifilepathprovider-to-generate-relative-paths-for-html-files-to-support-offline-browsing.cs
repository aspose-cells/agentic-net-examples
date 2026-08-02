// Title: Export Workbook to Separate HTML Files with Relative Paths Using IFilePathProvider (Aspose.Cells for .NET)
// Description: Demonstrates how to implement a custom IFilePathProvider that returns "{SheetName}.html" and configure HtmlSaveOptions (ExportActiveWorksheetOnly = false, IsFullPathLink = false) so each worksheet is saved as an individual HTML file linked relatively to the main workbook page, enabling offline browsing.
// Keywords: Aspose.Cells IFilePathProvider | C# HTML export relative links | Export workbook to multiple HTML files | .NET offline HTML workbook | HtmlSaveOptions IsFullPathLink false | Aspose.Cells custom file path provider | C# Aspose.Cells example | global developers | USA .NET community
// Common Searches: Aspose.Cells how to use IFilePathProvider for relative HTML links | Export each worksheet to its own HTML file in C# | Generate offline HTML workbook with Aspose.Cells | Relative paths in Aspose.Cells HTML export | C# save workbook as separate HTML pages
// Developer Intent: Create a multi‑sheet workbook export where every worksheet is saved as an individual HTML file linked via relative paths for portable, offline use.
// Use Cases: Produce a self‑contained HTML report that can be opened on any device without internet access. | Integrate the RelativePathProvider into a web service that generates downloadable workbook snapshots. | Customize file naming (e.g., add timestamps or subfolders) while preserving correct relative navigation between pages.
// AI Prompts: Show how to modify RelativePathProvider to place worksheet HTML files in a "Sheets" subfolder while keeping links functional. | Generate C# code that appends a timestamp to each sheet's HTML filename using IFilePathProvider. | Explain the effect of setting IsFullPathLink to false on hyperlink generation in Aspose.Cells HTML export.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider that returns relative file names
    // Demonstrates how to implement a custom IFilePathProvider that returns "{SheetName}.html" and configure HtmlSaveOptions (ExportActiveWorksheetOnly = false, IsFullPathLink = false) so each worksheet is saved as an individual HTML file linked relatively to the main workbook page, enabling offline browsing.
    public class RelativePathProvider : IFilePathProvider
    {
        // Returns a relative path for each worksheet based on its name
        public string GetFullName(string sheetName)
        {
            // Example: "Sheet1.html" – a relative file in the same folder as the main HTML file
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a workbook with multiple worksheets
            Workbook workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";
            workbook.Worksheets[0].Cells["A1"].PutValue("Data in Sheet1");

            // Add a second worksheet
            int sheetIndex = workbook.Worksheets.Add();
            workbook.Worksheets[sheetIndex].Name = "Sheet2";
            workbook.Worksheets[sheetIndex].Cells["A1"].PutValue("Data in Sheet2");

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                // Export all worksheets separately (not only the active one)
                ExportActiveWorksheetOnly = false,

                // Use relative links between the main HTML file and the sheet files
                IsFullPathLink = false,

                // Assign the custom file path provider
                FilePathProvider = new RelativePathProvider()
            };

            // Save the workbook to HTML. The main file will be "Workbook.html"
            // and each worksheet will be saved as "Sheet1.html", "Sheet2.html", etc.
            string outputPath = Path.Combine(Environment.CurrentDirectory, "Workbook.html");
            workbook.Save(outputPath, saveOptions);

            Console.WriteLine($"Workbook exported to HTML with relative links at: {outputPath}");
        }
    }
}
