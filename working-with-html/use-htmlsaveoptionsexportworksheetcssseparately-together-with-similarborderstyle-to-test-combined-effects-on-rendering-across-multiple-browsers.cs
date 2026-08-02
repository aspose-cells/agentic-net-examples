// Title: Export Worksheet CSS Separately & Apply Similar Border Style with Aspose.Cells HTML Save (C#)
// Description: This example creates a workbook with thick and thin bordered cells, then saves it to HTML using HtmlSaveOptions. It enables ExportWorksheetCSSSeparately to generate a distinct CSS file for each worksheet and activates ExportSimilarBorderStyle to provide fallback border rendering for browsers that lack native support. The HTML and CSS files are written to a desktop folder for easy cross‑browser testing.
// Keywords: Aspose.Cells | HtmlSaveOptions | ExportWorksheetCSSSeparately | ExportSimilarBorderStyle | C# HTML export | separate CSS per worksheet | border rendering browsers | cross‑browser HTML output | Aspose.Cells example
// Common Searches: Aspose.Cells export worksheet CSS to separate file | How to use ExportSimilarBorderStyle in Aspose.Cells | HTML export with per‑worksheet CSS C# | Test thick border rendering in Chrome Firefox Aspose.Cells | Set attached files directory for Aspose.Cells HTML save
// Developer Intent: Generate HTML from a workbook with individual CSS files per worksheet and a fallback border style to compare rendering across browsers.
// Use Cases: Produce modular HTML reports where each worksheet maintains its own stylesheet. | Ensure visual consistency on legacy browsers by falling back to similar border styles. | Automate placement of HTML and CSS assets in a known folder for manual or scripted UI testing.
// AI Prompts: Write C# code that saves an Aspose.Cells workbook to HTML with ExportWorksheetCSSSeparately = true and ExportSimilarBorderStyle = true, then open the output in Chrome and Firefox to compare border appearance. | Explain how ExportSimilarBorderStyle maps unsupported border types to CSS properties and why it improves cross‑browser compatibility. | Create a PowerShell script that checks the desktop folder for the generated HTML file and the separate CSS files after the Aspose.Cells export completes.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This example creates a workbook with thick and thin bordered cells, then saves it to HTML using HtmlSaveOptions. It enables ExportWorksheetCSSSeparately to generate a distinct CSS file for each worksheet and activates ExportSimilarBorderStyle to provide fallback border rendering for browsers that lack native support. The HTML and CSS files are written to a desktop folder for easy cross‑browser testing.
    public class ExportWorksheetCssAndSimilarBorderDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DemoSheet";

            // Add a cell with a thick border (may not be supported by all browsers)
            Cell thickCell = sheet.Cells["A1"];
            thickCell.PutValue("Thick Border");
            Style thickStyle = workbook.CreateStyle();
            thickStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
            thickStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
            thickCell.SetStyle(thickStyle);

            // Add a cell with a thin border (generally supported)
            Cell thinCell = sheet.Cells["B2"];
            thinCell.PutValue("Thin Border");
            Style thinStyle = workbook.CreateStyle();
            thinStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            thinStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            thinCell.SetStyle(thinStyle);

            // Configure HTML save options
            HtmlSaveOptions options = new HtmlSaveOptions
            {
                // Export CSS for each worksheet into a separate file
                ExportWorksheetCSSSeparately = true,
                // Use similar border style for browsers that don't support the original style
                ExportSimilarBorderStyle = true,
                // Define output directory (CSS files will be placed here)
                AttachedFilesDirectory = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "HtmlExport"),
                // Ensure the directory is created automatically if it does not exist
                CreateDirectory = true
            };

            // Ensure the output directory exists
            if (!Directory.Exists(options.AttachedFilesDirectory))
            {
                Directory.CreateDirectory(options.AttachedFilesDirectory);
            }

            // Save the workbook as HTML
            string htmlPath = Path.Combine(options.AttachedFilesDirectory, "Demo.html");
            workbook.Save(htmlPath, options);

            Console.WriteLine($"HTML file saved to: {htmlPath}");
            Console.WriteLine($"Separate CSS files are located in: {options.AttachedFilesDirectory}");
        }
    }
}
