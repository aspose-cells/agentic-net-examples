using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class HtmlToExcelWithSvg
{
    static void Main()
    {
        try
        {
            // Path to the source HTML file that contains embedded SVG graphics
            string htmlPath = "source.html";

            // Verify that the HTML file exists before loading
            if (!File.Exists(htmlPath))
            {
                Console.WriteLine($"HTML file not found: {htmlPath}");
                return;
            }

            // Load the HTML file into a workbook (Aspose.Cells parses the HTML content)
            Workbook workbook = new Workbook(htmlPath);

            // Path to the SVG file (extracted from the HTML)
            string svgPath = "image.svg";

            // Verify that the SVG file exists before reading
            if (!File.Exists(svgPath))
            {
                Console.WriteLine($"SVG file not found: {svgPath}");
                return;
            }

            // Read the SVG data into a byte array
            byte[] svgData = File.ReadAllBytes(svgPath);

            // Optional raster fallback for older Excel versions (null means no fallback)
            byte[] fallbackData = null;

            // Add the SVG as a scalable vector shape to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Position the SVG at row 5, column 3 (zero‑based indices)
            int topRow = 4;      // Row index (0‑based)
            int top = 0;         // Vertical offset in pixels
            int leftColumn = 2;  // Column index (0‑based)
            int left = 0;        // Horizontal offset in pixels
            int height = -1;     // -1 lets Excel auto‑size the shape height
            int width = -1;      // -1 lets Excel auto‑size the shape width

            shapes.AddSvg(topRow, top, leftColumn, left, height, width, svgData, fallbackData);

            // Save the workbook as an Excel file, preserving the SVG as a vector shape
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
            Console.WriteLine("Workbook saved successfully as output.xlsx");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}