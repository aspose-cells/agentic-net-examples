// Title: Perform pixel‑accurate hit testing on an Aspose.Cells worksheet to map mouse clicks to Excel cell addresses in C#
// AI Prompts: Generate C# code that builds cumulative column and row pixel offset lists from a Worksheet and returns the Excel cell reference for a given mouse X/Y coordinate using Aspose.Cells. | Create a binary‑search method that receives the pixel offset lists and a coordinate pair, finds the matching column and row indices, and converts the column index to its letter representation. | Demonstrate how to export a worksheet to SVG with Aspose.Cells for visual verification before applying the hit‑testing logic.
// Common Searches: Aspose.Cells C# convert mouse click position to Excel cell reference | pixel based hit test for worksheet cells using Aspose.Cells | retrieve column width in pixels and row height in pixels with Aspose.Cells | binary search column index from pixel coordinate Aspose.Cells example | export worksheet to SVG for hit testing Aspose.Cells C#
// Tags: pixel hit testing Aspose.Cells C# | worksheet column width pixel Aspose.Cells | worksheet row height pixel Aspose.Cells | binary search cell index Aspose.Cells | export worksheet to SVG Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsSvgHitTest
{
    // The sample loads an Excel workbook with Aspose.Cells, optionally renders the first worksheet to SVG for visual checks, builds cumulative pixel positions for each column and row using GetColumnWidthPixel and GetRowHeightPixel, and then performs hit testing by binary‑searching these lists to locate the column and row that contain a given mouse X/Y coordinate, finally returning the corresponding cell address such as "B3".
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Path to the input workbook
                string workbookPath = "input.xlsx";

                // Verify that the workbook file exists
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"Error: Workbook file \"{workbookPath}\" not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(workbookPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Export the worksheet to SVG (optional – for visual verification)
                string svgPath = "output.svg";
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    SaveFormat = SaveFormat.Svg,
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                };
                SheetRender render = new SheetRender(sheet, imgOptions);
                render.ToImage(0, svgPath);

                // Build cumulative pixel positions for columns and rows
                List<double> columnPositions = BuildColumnPositions(sheet);
                List<double> rowPositions = BuildRowPositions(sheet);

                // Example hit‑test coordinates (replace with actual mouse coordinates)
                double mouseX = 150; // pixel from left of the worksheet area
                double mouseY = 80;  // pixel from top of the worksheet area

                // Perform hit‑testing
                string cellAddress = HitTest(mouseX, mouseY, columnPositions, rowPositions);
                Console.WriteLine($"Mouse at ({mouseX}, {mouseY}) hits cell: {cellAddress ?? "outside range"}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        // Builds a list where each entry is the X‑pixel position of the left edge of a column.
        // Includes an extra entry for the rightmost edge (total width).
        private static List<double> BuildColumnPositions(Worksheet sheet)
        {
            List<double> positions = new List<double>();
            double current = 0.0;
            positions.Add(current); // left edge of column 0

            int totalCols = sheet.Cells.MaxColumn + 1; // cover all used columns
            for (int col = 0; col < totalCols; col++)
            {
                double widthPx = sheet.Cells.GetColumnWidthPixel(col);
                current += widthPx;
                positions.Add(current); // left edge of next column
            }
            return positions;
        }

        // Builds a list where each entry is the Y‑pixel position of the top edge of a row.
        // Includes an extra entry for the bottommost edge (total height).
        private static List<double> BuildRowPositions(Worksheet sheet)
        {
            List<double> positions = new List<double>();
            double current = 0.0;
            positions.Add(current); // top edge of row 0

            int totalRows = sheet.Cells.MaxRow + 1; // cover all used rows
            for (int row = 0; row < totalRows; row++)
            {
                double heightPx = sheet.Cells.GetRowHeightPixel(row);
                current += heightPx;
                positions.Add(current); // top edge of next row
            }
            return positions;
        }

        // Performs hit‑testing given a point (x, y) in pixel coordinates.
        // Returns the cell address (e.g., "B3") or null if the point is outside the used range.
        private static string HitTest(double x, double y, List<double> columnPositions, List<double> rowPositions)
        {
            int colIndex = FindIndex(columnPositions, x);
            int rowIndex = FindIndex(rowPositions, y);

            if (colIndex == -1 || rowIndex == -1)
                return null;

            string columnName = GetColumnName(colIndex);
            return $"{columnName}{rowIndex + 1}";
        }

        // Binary search to locate the interval that contains the coordinate.
        // Returns the zero‑based index of the interval, or -1 if out of bounds.
        private static int FindIndex(List<double> positions, double coordinate)
        {
            int low = 0;
            int high = positions.Count - 2; // last valid interval start index

            while (low <= high)
            {
                int mid = (low + high) / 2;
                if (coordinate >= positions[mid] && coordinate < positions[mid + 1])
                    return mid;
                if (coordinate < positions[mid])
                    high = mid - 1;
                else
                    low = mid + 1;
            }
            return -1; // outside range
        }

        // Converts a zero‑based column index to its Excel column letter(s).
        private static string GetColumnName(int index)
        {
            const int letters = 26;
            string name = string.Empty;
            int dividend = index + 1;

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % letters;
                name = Convert.ToChar('A' + modulo) + name;
                dividend = (dividend - modulo) / letters;
            }
            return name;
        }
    }
}
