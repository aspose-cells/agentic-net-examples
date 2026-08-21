// Title: Convert HTML with Embedded SVG to Excel Vector Shapes using Aspose.Cells C#
// Description: A C# example that reads an HTML file, extracts each <svg> element, converts the markup to a UTF‑8 byte array, and inserts the graphics into an Excel worksheet as scalable vector shapes via Aspose.Cells' AddSvg method. Non‑SVG content is placed in a cell and the workbook is saved as .xlsx.
// Keywords: Aspose.Cells AddSvg | HTML to Excel SVG conversion | C# extract SVG from HTML | embed vector graphics in Excel | convert SVG to Excel shape | .NET Aspose.Cells example
// Common Searches: how to import SVG from HTML into Excel using Aspose.Cells | C# extract <svg> tags and add as vector shapes in Excel | Aspose.Cells AddSvg multiple SVGs example | preserve SVG quality when converting HTML to XLSX | convert web page charts (SVG) to Excel workbook
// Developer Intent: Transform an HTML document containing embedded SVG graphics into an Excel workbook while keeping each SVG as a scalable vector shape.
// Use Cases: Generate Excel reports from web dashboards that use SVG charts, retaining crisp vector rendering. | Create product catalogs that import SVG icons from HTML into Excel for high‑resolution printing. | Migrate HTML‑based data visualizations to offline Excel files without losing vector quality.
// AI Prompts: Show a C# code snippet that reads an HTML file, extracts all <svg> elements, and adds them to an Aspose.Cells worksheet as vector shapes with automatic sizing. | Explain how to adjust position, offset, and scaling of SVG shapes inserted with the AddSvg method in Aspose.Cells. | Suggest ways to preserve inline CSS styles of SVG elements when converting them to Excel vector shapes.

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsHtmlToExcel
{
    // A C# example that reads an HTML file, extracts each <svg> element, converts the markup to a UTF‑8 byte array, and inserts the graphics into an Excel worksheet as scalable vector shapes via Aspose.Cells' AddSvg method. Non‑SVG content is placed in a cell and the workbook is saved as .xlsx.
    class Program
    {
        static void Main()
        {
            // Path to the source HTML file containing embedded SVG graphics
            string htmlPath = "input.html";

            // Path for the generated Excel workbook
            string excelPath = "output.xlsx";

            // Read the entire HTML content
            string htmlContent = File.ReadAllText(htmlPath, Encoding.UTF8);

            // Create a new workbook (empty)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Regular expression to capture <svg ...>...</svg> blocks (including attributes)
            string svgPattern = @"<svg[\s\S]*?<\/svg>";
            MatchCollection matches = Regex.Matches(htmlContent, svgPattern, RegexOptions.IgnoreCase);

            // Starting cell for placing SVGs (A1)
            int startRow = 0;
            int startCol = 0;

            // Iterate over each found SVG and add it as a scalable vector shape
            foreach (Match match in matches)
            {
                // Convert the SVG markup to UTF-8 byte array
                byte[] svgBytes = Encoding.UTF8.GetBytes(match.Value);

                // Add the SVG to the worksheet.
                // Height and width are set to -1 to let Excel auto‑size the shape.
                // Offsets (top, left) are set to 0 for simplicity.
                shapes.AddSvg(startRow, 0, startCol, 0, -1, -1, svgBytes, null);

                // Move to the next column for the next SVG (you can adjust layout as needed)
                startCol++;
                if (startCol >= 5) // after 5 columns, move to next row
                {
                    startCol = 0;
                    startRow++;
                }
            }

            // Optionally, you can also import the plain text part of the HTML into cells.
            // Here we simply place the raw HTML (without SVG tags) into cell A1.
            string htmlWithoutSvg = Regex.Replace(htmlContent, svgPattern, string.Empty, RegexOptions.IgnoreCase);
            sheet.Cells["A1"].PutValue(htmlWithoutSvg.Trim());

            // Save the workbook as an Excel file
            workbook.Save(excelPath, SaveFormat.Xlsx);

            Console.WriteLine($"HTML converted to Excel with {matches.Count} SVG shape(s) saved at '{excelPath}'.");
        }
    }
}
