// Title: C# – Verify Aspose.Cells ExportSimilarBorderStyle Double‑Border Fallback in Firefox
// Description: Creates a workbook, applies a blue double‑line border to cell A1, saves the sheet as HTML with ExportSimilarBorderStyle enabled and IsBorderCollapsed disabled, extracts the first <style> block, and shows how to open the file in Firefox to confirm the fallback border rendering.
// Keywords: Aspose.Cells | ExportSimilarBorderStyle | HTMLSaveOptions | double border fallback | Firefox CSS rendering | C# Excel to HTML | cell border style | CSS extraction | IsBorderCollapsed | Aspose.Cells HTML export
// Common Searches: Aspose.Cells ExportSimilarBorderStyle example C# | how to test double border rendering in Firefox with Aspose.Cells | extract generated CSS from Aspose.Cells HTML output | fallback border styles for unsupported Excel borders | C# save workbook as HTML with similar border style
// Developer Intent: Generate HTML from an Excel workbook with ExportSimilarBorderStyle enabled to inspect the CSS fallback for double borders, especially in Firefox.
// Use Cases: Confirm that double‑line borders are rendered using fallback CSS when the browser lacks native support. | Programmatically retrieve the <style> block from the exported HTML for automated validation. | Integrate the export into a CI pipeline that opens the HTML in Firefox and verifies visual appearance.
// AI Prompts: Write a unit test in C# that saves a workbook with ExportSimilarBorderStyle=true and asserts that the output CSS contains a fallback rule for double borders. | Provide a PowerShell script that launches Firefox headlessly, loads the generated HTML, and checks the computed style of the cell for the expected border width and color. | Explain the internal mapping Aspose.Cells uses to convert unsupported Excel border types (e.g., Double) to CSS properties when ExportSimilarBorderStyle is enabled.

using System;
using System.IO;
using System.Drawing;
using Aspose.Cells;

namespace AsposeCellsBorderFallbackDemo
{
    // Creates a workbook, applies a blue double‑line border to cell A1, saves the sheet as HTML with ExportSimilarBorderStyle enabled and IsBorderCollapsed disabled, extracts the first <style> block, and shows how to open the file in Firefox to confirm the fallback border rendering.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Prepare a style with a border type that is not widely supported by browsers (e.g., Double)
            Style doubleBorderStyle = workbook.CreateStyle();
            doubleBorderStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Double;
            doubleBorderStyle.Borders[BorderType.TopBorder].Color = Color.Blue;
            doubleBorderStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Double;
            doubleBorderStyle.Borders[BorderType.BottomBorder].Color = Color.Blue;
            doubleBorderStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Double;
            doubleBorderStyle.Borders[BorderType.LeftBorder].Color = Color.Blue;
            doubleBorderStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Double;
            doubleBorderStyle.Borders[BorderType.RightBorder].Color = Color.Blue;

            // Apply the style to a cell
            Cell cell = sheet.Cells["A1"];
            cell.PutValue("Double Border");
            cell.SetStyle(doubleBorderStyle);

            // Create HTML save options with ExportSimilarBorderStyle enabled
            HtmlSaveOptions htmlOptions = new HtmlSaveOptions(SaveFormat.Html)
            {
                ExportSimilarBorderStyle = true, // Enable fallback border rendering
                IsBorderCollapsed = false        // Keep borders separate for clearer CSS
            };

            // Define output HTML path
            string outputHtml = Path.Combine(Environment.CurrentDirectory, "BorderFallback.html");

            // Save the workbook as HTML using the specified options
            workbook.Save(outputHtml, htmlOptions);

            Console.WriteLine($"HTML file saved to: {outputHtml}");

            // Read the generated HTML and output the CSS block that defines the cell borders
            // This helps to inspect how Aspose.Cells rendered the fallback style.
            string htmlContent = File.ReadAllText(outputHtml);
            Console.WriteLine("\n--- Extracted CSS for cell borders ---\n");

            // Simple extraction: locate the first <style> block and print its content
            int styleStart = htmlContent.IndexOf("<style");
            if (styleStart >= 0)
            {
                int styleEnd = htmlContent.IndexOf("</style>", styleStart);
                if (styleEnd > styleStart)
                {
                    string styleBlock = htmlContent.Substring(styleStart, styleEnd - styleStart + 8);
                    Console.WriteLine(styleBlock);
                }
                else
                {
                    Console.WriteLine("No closing </style> tag found.");
                }
            }
            else
            {
                Console.WriteLine("No <style> block found in the generated HTML.");
            }

            // Note: Open the generated HTML file in Firefox to visually verify the fallback border rendering.
        }
    }
}
