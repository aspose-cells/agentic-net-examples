// Title: C# Example: Detect and Reposition Shapes Outside Worksheet Boundaries with Aspose.Cells
// Description: Loads an Excel workbook, scans each worksheet for shapes that lie beyond the used range, calculates the rows and columns a shape occupies, and automatically moves the shape inside the visible area before saving the file.
// Keywords: Aspose.Cells shape reposition | C# Excel shape boundary | move shapes within worksheet limits | prevent shape clipping Aspose.Cells | .NET Excel shape adjustment | shape overflow correction | Aspose.Cells example GitHub
// Common Searches: Aspose.Cells move shape inside worksheet | C# reposition Excel shapes that exceed sheet size | detect shapes outside used range Aspose.Cells | adjust shape location to avoid clipping in Excel | sample code for shape boundary correction Aspose.Cells
// Developer Intent: Identify shapes that extend past the worksheet edges and automatically shift them so the entire shape remains visible within the sheet.
// Use Cases: Fixing charts or images that are placed off‑sheet by automated generation tools. | Cleaning up workbooks before printing or converting to PDF to avoid cut‑off graphics. | Ensuring consistent layout when importing external Excel files that contain mis‑aligned shapes.
// AI Prompts: Create a reusable method that accepts a Worksheet and repositions any out‑of‑bounds shapes using Aspose.Cells. | Rewrite the shape‑size calculation to use pixel dimensions instead of row height and column width. | Explain how to handle shape repositioning when the worksheet contains merged cells or hidden rows.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeBoundaryAdjustment
{
    // Loads an Excel workbook, scans each worksheet for shapes that lie beyond the used range, calculates the rows and columns a shape occupies, and automatically moves the shape inside the visible area before saving the file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Input and output file paths
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the last used row and column indexes (zero‑based)
                    int maxRow = sheet.Cells.MaxRow;
                    int maxCol = sheet.Cells.MaxColumn;

                    // Iterate through each shape on the current worksheet
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Ensure the shape's top‑left corner is inside the worksheet bounds
                        if (shape.UpperLeftRow > maxRow)
                            shape.UpperLeftRow = maxRow;

                        if (shape.UpperLeftColumn > maxCol)
                            shape.UpperLeftColumn = maxCol;

                        // Approximate rows/columns occupied based on shape size
                        double rowHeight = sheet.Cells.GetRowHeight(shape.UpperLeftRow); // height in points
                        int rowsOccupied = (int)Math.Ceiling(shape.Height / rowHeight);

                        double columnWidth = sheet.Cells.GetColumnWidth(shape.UpperLeftColumn); // width in characters
                        // Convert column width (characters) to points using a typical conversion factor (approx. 7 points per character)
                        double columnWidthInPoints = columnWidth * 7.0;
                        int colsOccupied = (int)Math.Ceiling(shape.Width / columnWidthInPoints);

                        // Adjust if the shape would exceed the bottom edge
                        if (shape.UpperLeftRow + rowsOccupied - 1 > maxRow)
                        {
                            int newTopRow = Math.Max(0, maxRow - rowsOccupied + 1);
                            shape.UpperLeftRow = newTopRow;
                        }

                        // Adjust if the shape would exceed the right edge
                        if (shape.UpperLeftColumn + colsOccupied - 1 > maxCol)
                        {
                            int newLeftCol = Math.Max(0, maxCol - colsOccupied + 1);
                            shape.UpperLeftColumn = newLeftCol;
                        }
                    }
                }

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
