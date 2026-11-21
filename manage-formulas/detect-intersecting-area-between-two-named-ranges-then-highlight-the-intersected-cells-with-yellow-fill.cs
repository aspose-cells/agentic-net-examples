// Title: How to highlight the intersecting area of two named ranges with a yellow fill using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find overlapping cells of two named ranges | C# apply yellow background to intersected range in Excel using Aspose.Cells | How to get intersection of named ranges and style cells with Aspose.Cells .NET | Detect intersecting area of named ranges and set fill color with Aspose.Cells | Calculate CellArea overlap and highlight cells in Aspose.Cells for .NET
// Tags: named range intersection calculation Aspose.Cells | apply solid yellow fill style Aspose.Cells | CellArea overlap detection C# | highlight intersecting cells Excel Aspose.Cells | programmatic cell styling Aspose.Cells .NET

using Aspose.Cells;
using System;
using System.Drawing;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, retrieves two named ranges, converts them to CellArea objects, computes their overlapping region, creates a solid yellow style, applies the style to each cell in the intersected area, and saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the two named ranges (ensure the names exist in the workbook)
            AsposeRange range1 = workbook.Worksheets.GetRangeByName("Range1");
            AsposeRange range2 = workbook.Worksheets.GetRangeByName("Range2");

            if (range1 == null || range2 == null)
            {
                Console.WriteLine("One or both named ranges (Range1, Range2) were not found.");
                return;
            }

            // Convert ranges to CellArea objects using FirstRow/FirstColumn and RowCount/ColumnCount
            CellArea area1 = new CellArea
            {
                StartRow = range1.FirstRow,
                StartColumn = range1.FirstColumn,
                EndRow = range1.FirstRow + range1.RowCount - 1,
                EndColumn = range1.FirstColumn + range1.ColumnCount - 1
            };

            CellArea area2 = new CellArea
            {
                StartRow = range2.FirstRow,
                StartColumn = range2.FirstColumn,
                EndRow = range2.FirstRow + range2.RowCount - 1,
                EndColumn = range2.FirstColumn + range2.ColumnCount - 1
            };

            // Calculate the intersecting area
            CellArea intersect = new CellArea
            {
                StartRow = Math.Max(area1.StartRow, area2.StartRow),
                StartColumn = Math.Max(area1.StartColumn, area2.StartColumn),
                EndRow = Math.Min(area1.EndRow, area2.EndRow),
                EndColumn = Math.Min(area1.EndColumn, area2.EndColumn)
            };

            // Verify that an intersection actually exists
            if (intersect.StartRow <= intersect.EndRow && intersect.StartColumn <= intersect.EndColumn)
            {
                // Create a style with yellow fill
                Style yellowStyle = workbook.CreateStyle();
                yellowStyle.ForegroundColor = Color.Yellow;
                yellowStyle.Pattern = BackgroundType.Solid;

                // Assume both ranges are on the same worksheet; use the worksheet of the first range
                Worksheet sheet = range1.Worksheet;

                // Apply the yellow style to each cell in the intersected area
                for (int row = intersect.StartRow; row <= intersect.EndRow; row++)
                {
                    for (int col = intersect.StartColumn; col <= intersect.EndColumn; col++)
                    {
                        sheet.Cells[row, col].SetStyle(yellowStyle);
                    }
                }
            }
            else
            {
                Console.WriteLine("The named ranges do not intersect.");
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
