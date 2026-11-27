// Title: How to anchor a line shape’s connection points to cells A1 and B2 with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set line shape start and end cells A1 B2 | retrieve default connection points of a shape using Aspose.Cells | position a connector shape between two worksheet cells in Aspose.Cells | C# Aspose.Cells line connector anchored to specific cells | dynamic linking of Excel shapes to cell addresses with Aspose.Cells | how to get shape connection points array in Aspose.Cells .NET
// Tags: Aspose.Cells line shape anchoring | shape connection points Aspose.Cells | set shape UpperLeftRow UpperLeftColumn Aspose.Cells | Excel connector shape positioning .NET | C# shape anchor to specific cells Aspose.Cells | retrieve default shape connection points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace FlowConnectorDemo
{
    // The example creates a new workbook, adds a line shape, reads its default connection points, anchors the shape’s start and end to cells A1 and B2 by setting UpperLeftRow/Column and LowerRightRow/Column, and saves the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            const string outputPath = "FlowConnectorDemo.xlsx";

            try
            {
                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // ------------------------------------------------------------
                // 1. Add a line shape (as a simple connector) to the sheet
                // ------------------------------------------------------------
                // Use AddLine to create a line shape
                Shape line = worksheet.Shapes.AddLine(
                    0, // upper left row
                    0, // upper left column
                    0, // top offset (pixels)
                    0, // left offset (pixels)
                    0, // height (pixels) – will be adjusted later
                    0  // width (pixels) – will be adjusted later
                );

                // ------------------------------------------------------------
                // 2. Retrieve and display the default connection points of the shape
                // ------------------------------------------------------------
                float[][] connectionPoints = line.GetConnectionPoints();

                Console.WriteLine("Default connection points of the line shape:");
                for (int i = 0; i < connectionPoints.Length; i++)
                {
                    Console.WriteLine($"  Point {i + 1}: X = {connectionPoints[i][0]}, Y = {connectionPoints[i][1]}");
                }

                // ------------------------------------------------------------
                // 3. Position the line so that its ends align with cells A1 and B2
                // ------------------------------------------------------------
                int startRow = 0, startColumn = 0; // A1
                int endRow = 1, endColumn = 1;     // B2

                // Set the shape's anchor to span the two cells
                line.UpperLeftRow = startRow;
                line.UpperLeftColumn = startColumn;
                line.LowerRightRow = endRow;
                line.LowerRightColumn = endColumn;

                // ------------------------------------------------------------
                // 4. Verify the positioning
                // ------------------------------------------------------------
                Console.WriteLine($"Line positioned from cell {CellIndexToName(startRow, startColumn)} to {CellIndexToName(endRow, endColumn)}.");

                // ------------------------------------------------------------
                // 5. Save the workbook
                // ------------------------------------------------------------
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }

        private static string CellIndexToName(int row, int column)
        {
            // Convert column index to letters (e.g., 0 -> A)
            string colName = "";
            int dividend = column + 1;
            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                colName = Convert.ToChar('A' + modulo) + colName;
                dividend = (dividend - modulo) / 26;
            }

            // Row index is zero‑based; Excel rows start at 1
            return $"{colName}{row + 1}";
        }
    }
}
