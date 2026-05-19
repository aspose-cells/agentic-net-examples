using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class InsertSvgLogo
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook(); // new workbook
        // Workbook workbook = new Workbook("input.xlsx"); // uncomment to load

        // Load the SVG logo into a byte array
        byte[] svgData = File.ReadAllBytes("company_logo.svg");

        // Insert the SVG into every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Get the shape collection for the current sheet
            ShapeCollection shapes = sheet.Shapes;

            // Add the SVG at the top‑right corner.
            // Row 0, column 10 is used as an example; adjust column index as needed.
            // Height and width set to -1 let Excel determine the size automatically.
            Picture picture = shapes.AddSvg(
                topRow: 0,          // upper left row index
                top: 0,             // vertical offset in pixels
                leftColumn: 10,     // upper left column index (near right edge)
                left: 0,            // horizontal offset in pixels
                height: -1,         // auto height
                width: -1,          // auto width
                svgData: svgData,
                compatibleImageData: null);
        }

        // Save the workbook with the inserted logos
        workbook.Save("output_with_logo.xlsx");
    }
}