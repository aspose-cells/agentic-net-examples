// Title: Export Aspose.Cells Workbook to Separate HTML Files with Relative Paths Using IFilePathProvider (.NET)
// Description: Demonstrates a custom IFilePathProvider that returns sheet‑name based HTML file names, configures HtmlSaveOptions with IsFullPathLink = false, and saves each worksheet as an independent HTML file (e.g., Sheet1.html, Data.html) for portable offline browsing.
// Keywords: Aspose.Cells IFilePathProvider | relative HTML links | HtmlSaveOptions IsFullPathLink false | export workbook to multiple HTML files | C# Aspose.Cells HTML export | .NET offline HTML reports | custom file path provider
// Common Searches: Aspose.Cells generate relative HTML file names per worksheet | How to export Excel sheets to separate HTML files in C# | IFilePathProvider example for offline HTML output | Set IsFullPathLink false Aspose.Cells | Aspose.Cells HTML export without absolute paths
// Developer Intent: Create independent, locally viewable HTML files for each worksheet by customizing the file naming logic.
// Use Cases: Produce self‑contained HTML reports for each sheet that can be zipped and shared without internet access. | Integrate a relative path provider in a web service that delivers workbook exports as portable HTML files. | Automate generation of multi‑sheet documentation where each sheet appears as its own HTML page with relative navigation.
// AI Prompts: Show a C# implementation of IFilePathProvider that returns safe relative file names for Aspose.Cells HTML export. | Explain step‑by‑step how to configure HtmlSaveOptions to generate offline‑compatible HTML files for every worksheet. | Provide a complete Aspose.Cells example that saves a workbook to separate HTML files with relative links and describe each setting.

using System;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Custom implementation of IFilePathProvider that generates relative file names
    // Demonstrates a custom IFilePathProvider that returns sheet‑name based HTML file names, configures HtmlSaveOptions with IsFullPathLink = false, and saves each worksheet as an independent HTML file (e.g., Sheet1.html, Data.html) for portable offline browsing.
    public class RelativePathProvider : IFilePathProvider
    {
        // Returns a relative path for each worksheet HTML file (e.g., "Sheet1.html")
        public string GetFullName(string sheetName)
        {
            // Ensure the file name is safe and uses a .html extension
            return $"{sheetName}.html";
        }
    }

    class Program
    {
        static void Main()
        {
            // Create a new workbook and add some sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            sheet.Cells["A1"].PutValue("Hello");
            sheet.Cells["A2"].PutValue("World");

            // Add a second worksheet to demonstrate multiple files
            Worksheet sheet2 = workbook.Worksheets.Add("Data");
            sheet2.Cells["B1"].PutValue(123);
            sheet2.Cells["B2"].PutValue(456);

            // Configure HTML save options
            HtmlSaveOptions saveOptions = new HtmlSaveOptions();
            // Use relative links (default is false, set explicitly for clarity)
            saveOptions.IsFullPathLink = false;
            // Assign the custom relative path provider
            saveOptions.FilePathProvider = new RelativePathProvider();

            // Save the workbook; each worksheet will be exported to a separate HTML file
            // with relative links such as "Sheet1.html" and "Data.html"
            workbook.Save("Workbook.html", saveOptions);

            Console.WriteLine("Workbook exported to HTML with relative paths.");
        }
    }
}
