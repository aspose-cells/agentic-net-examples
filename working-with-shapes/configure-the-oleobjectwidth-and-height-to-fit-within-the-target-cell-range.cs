using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace OleObjectFitExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Define the target cell range where the OLE object should fit
                // Example: cells B2 (row 1, column 1) to D5 (row 4, column 3)
                int startRow = 1;      // zero‑based index
                int endRow = 4;
                int startColumn = 1;
                int endColumn = 3;

                // Prepare dummy OLE data (an empty byte array for demonstration)
                byte[] oleData = new byte[0];

                // Add the OLE object at the upper‑left cell of the range with temporary size
                int oleIndex = sheet.OleObjects.Add(startRow, startColumn, 10, 10, oleData);
                OleObject ole = sheet.OleObjects[oleIndex];

                // Calculate total width in pixels by summing column widths
                double totalWidthPixels = 0;
                for (int col = startColumn; col <= endColumn; col++)
                {
                    // Column width is returned in characters; convert to pixels (approx. 7 pixels per character)
                    double colWidthChars = sheet.Cells.GetColumnWidth(col);
                    totalWidthPixels += colWidthChars * 7;
                }

                // Calculate total height in pixels by summing row heights
                double totalHeightPixels = 0;
                for (int row = startRow; row <= endRow; row++)
                {
                    // Row height is returned in points; convert to pixels (1 point = 1/72 inch, 96 DPI)
                    double rowHeightPoints = sheet.Cells.GetRowHeight(row);
                    totalHeightPixels += rowHeightPoints * 96.0 / 72.0;
                }

                // Apply the calculated dimensions to the OLE object
                ole.Width = (int)Math.Round(totalWidthPixels);
                ole.Height = (int)Math.Round(totalHeightPixels);

                // Optional: make the object move and size with cells
                ole.Placement = PlacementType.MoveAndSize;

                // Ensure output directory exists
                string outputPath = "OleObjectFitWithinRange.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}