// Title: Export Excel with Double Borders to HTML Using ExportSimilarBorderStyle – Safari Fallback Test (C#)
// Description: C# example that creates a workbook, applies a blue double border to cells A1:D4, enables the HtmlSaveOptions ExportSimilarBorderStyle flag, and saves the file as SafariBorderFallback.html to verify how Safari handles unsupported double‑border styles.
// Keywords: Aspose.Cells | ExportSimilarBorderStyle | HTML export | double border | Safari fallback | cell border rendering | C# Aspose.Cells example | HtmlSaveOptions | cross‑browser border compatibility | Aspose.Cells Safari test
// Common Searches: Aspose.Cells ExportSimilarBorderStyle Safari | how to test double border rendering in Safari | HTML export with similar border style option | C# export Excel to HTML with border fallback | Aspose.Cells border compatibility across browsers
// Developer Intent: Generate an HTML file from a workbook that contains double borders while enabling ExportSimilarBorderStyle to observe the fallback rendering behavior in Safari.
// Use Cases: Validate visual consistency of double borders in browsers that lack native support, such as Safari. | Create HTML reports that gracefully degrade unsupported border styles using the similar‑border fallback. | Automate regression tests for cell border rendering after library upgrades or configuration changes.
// AI Prompts: Write C# code that parses the saved HTML and confirms whether double borders were replaced with a supported style for Safari. | Explain the internal algorithm Aspose.Cells uses to map unsupported border styles to similar ones during HTML export. | Recommend additional HtmlSaveOptions settings that improve cross‑browser cell border rendering and overall HTML quality.

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsSafariBorderTest
{
    // C# example that creates a workbook, applies a blue double border to cells A1:D4, enables the HtmlSaveOptions ExportSimilarBorderStyle flag, and saves the file as SafariBorderFallback.html to verify how Safari handles unsupported double‑border styles.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Prepare a style with a double border (may not be supported by all browsers)
                Style doubleBorderStyle = workbook.CreateStyle();
                doubleBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.TopBorder].Color = Color.Blue;
                doubleBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
                doubleBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;
                doubleBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
                doubleBorderStyle.Borders[BorderType.RightBorder].Color = Color.Blue;

                // Apply the style to a range of cells
                AsposeRange range = sheet.Cells.CreateRange("A1:D4");
                // Put the same value into all cells of the range
                range.PutValue("Double Border", false, false);
                range.ApplyStyle(doubleBorderStyle, new StyleFlag { Borders = true });

                // Create HTML save options and enable ExportSimilarBorderStyle
                HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
                {
                    ExportSimilarBorderStyle = true // Fallback to similar border style for unsupported browsers
                };

                // Define output path and ensure its directory exists
                string outputPath = "SafariBorderFallback.html";
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook as HTML
                workbook.Save(outputPath, htmlOptions);

                Console.WriteLine($"HTML file saved to {outputPath} with ExportSimilarBorderStyle enabled.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
