// Title: Adjust Excel column widths to the pixel width of the longest cell value (including formulas) using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells, iterates each column, computes the pixel width of every cell's displayed text using a fixed pixel‑per‑character factor plus padding, evaluates formulas when present, and applies SetColumnWidthPixel to resize the column. | Create a reusable C# method that accepts a Worksheet object, automatically determines the maximum pixel width of cell contents in each column (handling formula calculation), and updates all column widths with a configurable padding value.
// Common Searches: Aspose.Cells C# set column width based on longest cell text pixel measurement | How to auto‑size Excel columns by pixel width with Aspose.Cells .NET | Calculate column width from cell string length using Aspose.Cells SetColumnWidthPixel | Resize columns after evaluating formulas with Aspose.Cells in C# | C# adjust Excel column widths dynamically using pixel‑per‑character estimate
// Tags: SetColumnWidthPixel column auto‑fit Aspose.Cells | pixel‑per‑character column width calculation .NET | dynamic column resizing Excel Aspose.Cells | formula evaluation before column width adjustment | auto‑fit columns by pixel measurement C#

using System;
using System.IO;
using Aspose.Cells;

namespace AdjustColumnWidthsApp
{
    // The program loads Input.xlsx with Aspose.Cells, walks through each column to find the maximum pixel width of the displayed cell text (evaluating formulas when needed) using a 7‑pixel‑per‑character estimate plus padding, sets each column's width via SetColumnWidthPixel, and saves the adjusted workbook as Output.xlsx.
    class AdjustColumnWidths
    {
        static void Main()
        {
            try
            {
                const string inputPath = "Input.xlsx";
                const string outputPath = "Output.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet sheet = workbook.Worksheets[0];

                // Approximate pixel width per character for default font (Arial 10pt)
                const int pixelPerChar = 7;
                const int paddingPixels = 5;

                // Prepare calculation options for formula evaluation
                CalculationOptions calcOptions = new CalculationOptions();

                // Determine the used range of the worksheet
                int maxRow = sheet.Cells.MaxDataRow;
                int maxCol = sheet.Cells.MaxDataColumn;

                // Iterate through each column
                for (int col = 0; col <= maxCol; col++)
                {
                    int maxPixelWidth = 0;

                    // Iterate through each row in the current column
                    for (int row = 0; row <= maxRow; row++)
                    {
                        Cell cell = sheet.Cells[row, col];
                        if (cell == null) continue;

                        // Get the displayed string value
                        string cellText = cell.StringValue;

                        // If the cell contains a formula, calculate it first
                        if (cell.IsFormula)
                        {
                            cell.Calculate(calcOptions);
                            cellText = cell.StringValue;
                        }

                        // Estimate pixel width of the text
                        int textPixelWidth = (cellText?.Length ?? 0) * pixelPerChar + paddingPixels;

                        // Keep the maximum width for this column
                        if (textPixelWidth > maxPixelWidth)
                            maxPixelWidth = textPixelWidth;
                    }

                    // Apply the calculated width to the column (if any content was found)
                    if (maxPixelWidth > 0)
                        sheet.Cells.SetColumnWidthPixel(col, maxPixelWidth);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
