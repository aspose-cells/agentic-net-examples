// Title: Export a Single‑Sheet Workbook to HTML with a Separate CSS Folder using Aspose.Cells for .NET
// Description: Creates a desktop folder, builds a one‑worksheet workbook, adds sample data, and saves it as HTML with HtmlSaveOptions.ExportWorksheetCSSSeparately enabled. Aspose.Cells automatically generates a sub‑folder that contains the worksheet's CSS file.
// Keywords: Aspose.Cells HTML export | ExportWorksheetCSSSeparately | CreateDirectory option | C# HTML save options | .NET workbook to HTML | separate CSS folder | Aspose.Cells CSS assets
// Common Searches: Aspose.Cells export HTML separate CSS folder | How to use ExportWorksheetCSSSeparately in C# | Save Excel as HTML with external CSS using Aspose.Cells | Create CSS sub‑directory when exporting workbook to HTML | Aspose.Cells HtmlSaveOptions CreateDirectory example
// Developer Intent: Generate HTML from an Excel workbook while placing the worksheet's CSS in its own directory.
// Use Cases: Web‑ready reporting where styles are cached separately from HTML. | Batch conversion pipelines that need modular CSS per worksheet. | Desktop utilities that organize exported HTML and CSS assets on the user's desktop.
// AI Prompts: Generate C# code to export a multi‑sheet workbook to HTML with ExportWorksheetCSSSeparately and a custom CSS folder name. | Show how to post‑process the CSS files created by Aspose.Cells after HTML export. | Add robust error handling for missing output paths when using HtmlSaveOptions with CreateDirectory.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHtmlExport
{
    // Creates a desktop folder, builds a one‑worksheet workbook, adds sample data, and saves it as HTML with HtmlSaveOptions.ExportWorksheetCSSSeparately enabled. Aspose.Cells automatically generates a sub‑folder that contains the worksheet's CSS file.
    class Program
    {
        static void Main()
        {
            // Define output directory (desktop) and ensure it exists
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputDir = Path.Combine(desktopPath, "HtmlExport");
            Directory.CreateDirectory(outputDir);

            // Path for the main HTML file
            string htmlFilePath = Path.Combine(outputDir, "Workbook.html");

            // Create a new workbook with a single worksheet and add some data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            sheet.Cells["A1"].PutValue("Hello, Aspose.Cells!");
            sheet.Cells["B2"].PutValue(12345);

            // Configure HTML save options to export worksheet CSS separately
            HtmlSaveOptions saveOptions = new HtmlSaveOptions
            {
                ExportWorksheetCSSSeparately = true, // enable separate CSS files per worksheet
                CreateDirectory = true               // auto‑create directories if they don't exist
            };

            // Save the workbook as HTML; Aspose.Cells will create a CSS folder alongside the HTML file
            workbook.Save(htmlFilePath, saveOptions);

            Console.WriteLine($"HTML exported to: {htmlFilePath}");
            Console.WriteLine("Separate CSS files are stored in a sub‑folder next to the HTML file.");
        }
    }
}
