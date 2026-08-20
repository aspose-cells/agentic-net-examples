// Title: Export Excel to HTML with Separate CSS and Fallback Border Styles using Aspose.Cells for .NET
// Description: Shows how to save a workbook as HTML, write each worksheet’s styles to an external CSS file, and automatically replace unsupported border types with similar ones. Includes configuration of a custom cell‑class prefix.
// Keywords: Aspose.Cells HTML export | ExportWorksheetCSSSeparately | ExportSimilarBorderStyle | C# separate CSS file | border style fallback | cell CSS prefix | cross‑browser Excel to HTML | Aspose.Cells .NET example
// Common Searches: Aspose.Cells export HTML separate CSS | how to use ExportSimilarBorderStyle | C# save workbook as HTML with external stylesheet | fallback for thick borders in HTML output | custom CSS class prefix Aspose.Cells
// Developer Intent: Generate an HTML file from a workbook where the worksheet’s CSS is stored in its own file and any border style unsupported by the browser is substituted with a compatible alternative.
// Use Cases: Build web‑based reports that keep styling separate for caching and easier maintenance. | Maintain consistent border appearance across browsers by falling back to supported styles. | Integrate the HTML output into existing sites using a custom CSS class prefix to avoid naming conflicts.
// AI Prompts: Add code to export workbook images to a dedicated folder while keeping the CSS file separate. | Provide C# that reads the generated CSS file and injects extra rules for header rows. | Explain the internal mapping logic of ExportSimilarBorderStyle and list which border types are considered similar.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to save a workbook as HTML, write each worksheet’s styles to an external CSS file, and automatically replace unsupported border types with similar ones. Includes configuration of a custom cell‑class prefix.
    public class ExportWorksheetCssAndSimilarBorderDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data
                sheet.Cells["A1"].PutValue("Border Test");
                sheet.Cells["B1"].PutValue("Standard Border");
                sheet.Cells["C1"].PutValue("Thick Border");

                // Apply a standard thin border to B1
                Style thinStyle = workbook.CreateStyle();
                thinStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
                thinStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
                thinStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
                thinStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
                sheet.Cells["B1"].SetStyle(thinStyle);

                // Apply a thick border to C1 (border style not supported by some browsers)
                Style thickStyle = workbook.CreateStyle();
                thickStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thick;
                thickStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thick;
                thickStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thick;
                thickStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thick;
                sheet.Cells["C1"].SetStyle(thickStyle);

                // Configure HTML save options
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions
                {
                    // Export the worksheet CSS into a separate file (e.g., sheet0.css)
                    ExportWorksheetCSSSeparately = true,

                    // When a border style is not supported by the browser, export a similar style
                    ExportSimilarBorderStyle = true,

                    // Optional: give the CSS file a custom prefix for cell classes
                    CellCssPrefix = "cell-"
                };

                // Define output folder and file names
                string outputFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "HtmlExportDemo");

                // Ensure the output directory exists
                Directory.CreateDirectory(outputFolder);

                // Save the workbook as HTML; separate CSS files will be placed in the same folder
                string htmlPath = Path.Combine(outputFolder, "ExportDemo.html");
                workbook.Save(htmlPath, htmlOptions);

                Console.WriteLine($"HTML file saved to: {htmlPath}");
                Console.WriteLine("Check the output folder for the separate CSS file (e.g., sheet0.css).");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred during the export process:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportWorksheetCssAndSimilarBorderDemo.Run();
        }
    }
}
