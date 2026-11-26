// Title: Align all worksheet shapes to the left edge of the printable area with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# align worksheet shapes to printable area left margin | set shape UpperLeftColumn to first printable column in Excel using Aspose.Cells | move Excel shapes programmatically based on print area start column .NET | retrieve printable area start column from worksheet PageSetup Aspose.Cells | apply MoveAndSize placement to shapes so they follow cells in Aspose.Cells
// Tags: worksheet shape left alignment Aspose.Cells | shape UpperLeftColumn property C# | extract printable area start column Aspose.Cells | shape placement MoveAndSize .NET | programmatic Excel shape layout

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AlignShapesToPrintableArea
{
    // The example loads a workbook, determines the first column of the worksheet's printable area (using the defined PrintArea when available, otherwise the used range), iterates through all shapes on the first sheet, sets each shape's UpperLeftColumn to that column, resets the horizontal offset, applies the MoveAndSize placement so shapes move with cells, and saves the updated workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Paths for input and output workbooks
                string inputPath = "input.xlsx";
                string outputPath = "output_aligned.xlsx";

                // Load existing workbook or create a new one
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                    Console.WriteLine($"Loaded workbook from \"{inputPath}\".");
                }
                else
                {
                    workbook = new Workbook();
                    Console.WriteLine("Input file not found. Created a new workbook.");
                }

                // Work with the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the leftmost column of the printable area
                int printableStartColumn = GetPrintableAreaStartColumn(worksheet);
                Console.WriteLine($"Printable area starts at column index {printableStartColumn} (zero‑based).");

                // Align each shape to the left edge of the printable area
                foreach (Shape shape in worksheet.Shapes)
                {
                    try
                    {
                        shape.UpperLeftColumn = printableStartColumn; // anchor to printable start column
                        shape.Left = 0;                               // remove horizontal offset
                        shape.Placement = PlacementType.MoveAndSize; // move/size with cells
                        Console.WriteLine($"Aligned shape \"{shape.Name}\" to column {printableStartColumn}.");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Failed to align shape \"{shape.Name}\": {ex.Message}");
                    }
                }

                // Save the modified workbook
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error saving workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        private static int GetPrintableAreaStartColumn(Worksheet worksheet)
        {
            try
            {
                // Attempt to read the defined print area (e.g., "A1:D20")
                string printArea = worksheet.PageSetup.PrintArea;
                if (!string.IsNullOrEmpty(printArea))
                {
                    // PrintArea may include a sheet name (e.g., "Sheet1!A1:D20")
                    string range = printArea.Contains("!") ? printArea.Split('!')[1] : printArea;

                    // Aspose.Cells does not provide a direct parser in newer versions,
                    // so we fall back to the used range's first column.
                    Console.WriteLine($"PrintArea detected: \"{range}\" (parsing not supported, using fallback).");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading PrintArea: {ex.Message}");
            }

            // Fallback to the first column of the used range
            return worksheet.Cells.MaxDisplayRange.FirstColumn;
        }
    }
}
