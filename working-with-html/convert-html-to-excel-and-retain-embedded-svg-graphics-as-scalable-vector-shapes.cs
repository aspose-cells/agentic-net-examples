// Title: C# – Convert HTML to Excel and embed SVG graphics as scalable vector shapes with Aspose.Cells
// Description: Reads an HTML file, extracts every <svg> block with a regular expression, converts each markup to a UTF‑8 byte array, and inserts it into the first worksheet as a vector shape via Shapes.AddSvg. Shapes are auto‑sized and spaced before the workbook is saved as XLSX.
// Keywords: Aspose.Cells | AddSvg | SVG to Excel | HTML to XLSX | C# example | vector graphics in Excel | embed SVG | Excel shape API | Aspose.Cells for .NET | convert HTML to Excel
// Common Searches: How to add SVG images to Excel using Aspose.Cells C# | Extract SVG tags from HTML and place them in an XLSX file | Aspose.Cells Shapes.AddSvg usage example | Convert HTML page with SVG charts to Excel workbook | C# code to import vector graphics from HTML into Excel
// Developer Intent: Pull SVG elements from an HTML document and insert them as scalable vector shapes into an Excel workbook using Aspose.Cells.
// Use Cases: Automated reporting that preserves web‑based SVG charts in Excel for further analysis. | Archiving marketing emails containing SVG icons as high‑quality Excel assets. | Batch conversion of multiple HTML files, extracting each SVG and placing them in separate rows of a spreadsheet.
// AI Prompts: Write a reusable C# method that takes an HTML string and a worksheet, extracts all <svg> elements, and adds them with configurable row offset using Shapes.AddSvg. | Show how to adjust the width, height, and position of SVG shapes after they are added with Aspose.Cells. | Provide error‑handling code for cases where the input HTML contains no SVG tags but a valid Excel file must still be generated.

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Reads an HTML file, extracts every <svg> block with a regular expression, converts each markup to a UTF‑8 byte array, and inserts it into the first worksheet as a vector shape via Shapes.AddSvg. Shapes are auto‑sized and spaced before the workbook is saved as XLSX.
class HtmlToExcelWithSvg
{
    static void Main()
    {
        // Paths for input HTML and output Excel files
        string htmlPath = "input.html";
        string excelPath = "output.xlsx";

        // Load the entire HTML content
        string htmlContent = File.ReadAllText(htmlPath);

        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        ShapeCollection shapes = sheet.Shapes;

        // Regex to capture <svg ...>...</svg> blocks (including attributes and inner content)
        Regex svgRegex = new Regex(@"<svg[\s\S]*?<\/svg>", RegexOptions.IgnoreCase);
        MatchCollection matches = svgRegex.Matches(htmlContent);

        // Starting cell for placing SVGs (A1)
        int startRow = 0;
        int startColumn = 0;

        foreach (Match match in matches)
        {
            // Convert the SVG markup to UTF-8 byte array
            byte[] svgBytes = Encoding.UTF8.GetBytes(match.Value);

            // Add the SVG as a scalable vector shape.
            // Height and width are set to -1 to let Aspose.Cells calculate them automatically.
            // Offsets (top, left) are set to 0.
            shapes.AddSvg(startRow, 0, startColumn, 0, -1, -1, svgBytes, null);

            // Move to the next row for the next SVG (optional layout logic)
            startRow += 10; // Adjust spacing as needed
        }

        // Save the workbook as an Excel file
        workbook.Save(excelPath, SaveFormat.Xlsx);
    }
}
