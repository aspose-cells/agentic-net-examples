// Title: Check that Excel cell comments (tooltips) appear only when the cell text is truncated using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that walks through every worksheet and cell, finds cells containing comments, and reports those where the visible text fits within the column width. | Implement a C# helper method that decides if a cell's displayed string is truncated based on column width and wrap settings in Aspose.Cells, then use it to validate tooltip placement.
// Common Searches: Aspose.Cells how to determine if cell text is cut off by column width in C# | C# verify Excel comment appears only on truncated cells using Aspose.Cells | detect unnecessary tooltips in Excel worksheets with Aspose.Cells .NET | check if cell text is wrapped or truncated before showing tooltip Aspose.Cells | iterate used range and validate comment placement Aspose.Cells C#
// Tags: cell text truncation detection Aspose.Cells | validate comment tooltip based on column width | check wrap text setting Aspose.Cells | iterate worksheet used range .NET | column width character units Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace TooltipVerification
{
    // The example loads an Excel workbook with Aspose.Cells, iterates all worksheets and cells in the used range, and for each cell that has a comment (used as a tooltip) it calls a helper method that compares the cell's string length to the column width (ignoring wrapped cells) to decide if the text is truncated. Cells with comments but non‑truncated text are logged for correction.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                const string inputPath = "input.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    Aspose.Cells.Range usedRange = sheet.Cells.MaxDisplayRange;

                    // If the sheet is empty, skip it
                    if (usedRange == null)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Loop through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Check if the cell has a comment (used as tooltip)
                            if (cell.Comment != null)
                            {
                                // Determine if the cell's displayed text is truncated
                                bool isTruncated = IsCellTextTruncated(sheet, cell);

                                // If the text is NOT truncated but a tooltip exists, report it
                                if (!isTruncated)
                                {
                                    Console.WriteLine($"Worksheet \"{sheet.Name}\", Cell {cell.Name}: " +
                                                      "has a tooltip but its text is not truncated.");
                                }
                            }
                        }
                    }
                }

                // Optional: Save the workbook if modifications were made
                // workbook.Save("output.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }

        /// <param name="sheet">The worksheet containing the cell.</param>
        /// <param name="cell">The cell to evaluate.</param>
        /// <returns>True if the text is considered truncated; otherwise, false.</returns>
        private static bool IsCellTextTruncated(Worksheet sheet, Cell cell)
        {
            try
            {
                // Get the column width in Excel's character units
                double columnWidthInChars = sheet.Cells.GetColumnWidth(cell.Column);

                // Get the cell's displayed string (ignoring formulas)
                string cellText = cell.StringValue ?? string.Empty;

                // Empty cells cannot be truncated
                if (string.IsNullOrEmpty(cellText))
                    return false;

                // Check if the cell is set to wrap text; wrapped text is not considered truncated
                bool isWrapped = cell.GetStyle().IsTextWrapped;
                if (isWrapped)
                    return false;

                // Simple heuristic: treat as truncated when text length exceeds column width
                return cellText.Length > Math.Floor(columnWidthInChars);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error evaluating truncation for cell {cell.Name}: {ex.Message}");
                return false;
            }
        }
    }
}
