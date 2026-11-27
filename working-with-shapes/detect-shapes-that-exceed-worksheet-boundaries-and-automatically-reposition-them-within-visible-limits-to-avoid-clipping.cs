// Title: Detect and Reposition Shapes That Exceed Worksheet Boundaries in an Excel File Using Aspose.Cells for .NET (C#)
// Common Searches: how to find shapes that are outside the used range of an Excel worksheet with Aspose.Cells C# | C# code to move Excel shapes back inside visible rows and columns using Aspose.Cells | Aspose.Cells example for correcting shape clipping caused by out‑of‑bounds anchors | reposition rectangle shape that starts at row 15 column 8 in a .xlsx with Aspose.Cells | automatically adjust shape positions to fit within worksheet limits in .NET
// Tags: out-of-bounds shape detection Aspose.Cells C# | shape repositioning within worksheet limits | adjust Excel shape anchors programmatically | prevent shape clipping in .xlsx using Aspose.Cells | move rectangle shape to visible area C# | worksheet used range shape bounds handling

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace ShapeBoundaryReposition
{
    // The sample creates an Excel workbook, adds a rectangle shape placed beyond the data area, then loads the file, calculates the effective row and column limits from the used range and worksheet size, iterates each shape, and automatically shifts any shape whose start or end anchors lie outside those limits back into the visible worksheet area before saving the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths for the sample and result workbooks
            string samplePath = "sample.xlsx";
            string resultPath = "sample_repositioned.xlsx";

            try
            {
                // -----------------------------------------------------------------
                // STEP 1: Create a sample workbook with data and a shape that exceeds
                // the used range of the worksheet.
                // -----------------------------------------------------------------
                CreateSampleWorkbook(samplePath);
                Console.WriteLine($"Sample workbook created at '{samplePath}'.");

                // -----------------------------------------------------------------
                // STEP 2: Load the workbook (if it exists) and reposition any out‑of‑bounds shapes.
                // -----------------------------------------------------------------
                if (!File.Exists(samplePath))
                {
                    Console.WriteLine($"Error: The file '{samplePath}' was not found.");
                    return;
                }

                Workbook wb = new Workbook(samplePath);
                foreach (Worksheet ws in wb.Worksheets)
                {
                    RepositionOutOfBoundsShapes(ws);
                }

                // Save the modified workbook
                wb.Save(resultPath);
                Console.WriteLine($"Repositioned workbook saved at '{resultPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }

        private static void CreateSampleWorkbook(string filePath)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook wb = new Workbook();
                Worksheet ws = wb.Worksheets[0];

                // Populate some data (rows 0‑9, columns 0‑4)
                for (int row = 0; row < 10; row++)
                {
                    for (int col = 0; col < 5; col++)
                    {
                        ws.Cells[row, col].PutValue($"R{row}C{col}");
                    }
                }

                // Add a rectangle shape that is deliberately placed beyond the used range
                // Upper left corner at row 15, column 8 (outside the data area)
                // Height and width are arbitrary (100 pixels each)
                ShapeCollection shapes = ws.Shapes;
                shapes.AddRectangle(15, 0, 8, 0, 100, 100);

                // Save the workbook
                wb.Save(filePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to create sample workbook: {ex.Message}");
                throw;
            }
        }

        private static void RepositionOutOfBoundsShapes(Worksheet ws)
        {
            try
            {
                // Determine the used range limits.
                // MaxDataRow/Column return -1 when the sheet is empty; treat that as 0.
                int maxDataRow = ws.Cells.MaxDataRow;
                int maxDataColumn = ws.Cells.MaxDataColumn;

                if (maxDataRow < 0) maxDataRow = 0;
                if (maxDataColumn < 0) maxDataColumn = 0;

                // Absolute limits of the worksheet (total rows/columns - 1)
                int absoluteMaxRow = ws.Cells.Rows.Count - 1;
                int absoluteMaxColumn = ws.Cells.Columns.Count - 1;

                // Effective limits are the larger of used range and absolute limits.
                int limitRow = Math.Max(maxDataRow, absoluteMaxRow);
                int limitColumn = Math.Max(maxDataColumn, absoluteMaxColumn);

                ShapeCollection shapes = ws.Shapes;
                foreach (Shape shape in shapes)
                {
                    // Current anchor positions
                    int startRow = shape.UpperLeftRow;
                    int startColumn = shape.UpperLeftColumn;

                    // The shape's lower‑right corner (anchor positions)
                    int endRow = shape.LowerRightRow;
                    int endColumn = shape.LowerRightColumn;

                    bool moved = false;

                    // If the shape starts beyond the limits, shift it left/up.
                    if (startRow > limitRow)
                    {
                        shape.UpperLeftRow = limitRow;
                        moved = true;
                    }
                    if (startColumn > limitColumn)
                    {
                        shape.UpperLeftColumn = limitColumn;
                        moved = true;
                    }

                    // Re‑evaluate after possible start‑position adjustment
                    endRow = shape.LowerRightRow;
                    endColumn = shape.LowerRightColumn;

                    // If the shape's lower‑right corner exceeds limits, move it up/left.
                    if (endRow > limitRow)
                    {
                        int deltaRows = endRow - limitRow;
                        shape.UpperLeftRow = Math.Max(0, shape.UpperLeftRow - deltaRows);
                        moved = true;
                    }
                    if (endColumn > limitColumn)
                    {
                        int deltaCols = endColumn - limitColumn;
                        shape.UpperLeftColumn = Math.Max(0, shape.UpperLeftColumn - deltaCols);
                        moved = true;
                    }

                    if (moved)
                    {
                        Console.WriteLine($"Shape (type: {shape.GetType().Name}) repositioned to Row {shape.UpperLeftRow}, Column {shape.UpperLeftColumn}.");
                    }
                    else
                    {
                        Console.WriteLine($"Shape (type: {shape.GetType().Name}) is within bounds.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while repositioning shapes: {ex.Message}");
            }
        }
    }
}
