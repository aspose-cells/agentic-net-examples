// Title: C# Utility to Detect Formatting‑Only Cells in an Aspose.Cells Worksheet
// Description: A concise C# method that scans the used range of an Aspose.Cells worksheet, identifies cells where a style is applied (IsStyleSet) but the cell contains no value and is not a formula, and returns true as soon as such a formatting‑only cell is found. Ideal for auditing Excel templates and optimizing workbook size.
// Keywords: Aspose.Cells formatting only cells | C# detect empty styled cells | worksheet IsStyleSet check | Excel template validation .NET | remove style‑only cells Aspose
// Common Searches: how to find cells with only formatting using Aspose.Cells C# | detect empty cells that have a style applied in a worksheet | C# method to check for formatting‑only cells in Excel | Aspose.Cells identify cells with style but no data
// Developer Intent: Determine whether any cell in a worksheet has a style applied while lacking a value or formula.
// Use Cases: Validate Excel templates before distribution to ensure no stray formatting‑only cells remain. | Identify and clean up style‑only cells to reduce file size and improve performance. | Flag worksheets that may cause rendering delays due to excessive formatting without content.
// AI Prompts: Generate unit tests for HasFormattingOnlyInitializedCells covering styled empty cells, styled cells with data, and formula cells. | Rewrite the method to use Cells.MaxDataRow and Cells.MaxDataColumn for a more efficient scan. | Create a version that returns a list of addresses for all formatting‑only cells instead of a boolean.

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetUtilitiesDemo
{
    // A concise C# method that scans the used range of an Aspose.Cells worksheet, identifies cells where a style is applied (IsStyleSet) but the cell contains no value and is not a formula, and returns true as soon as such a formatting‑only cell is found. Ideal for auditing Excel templates and optimizing workbook size.
    public static class WorksheetUtilities
    {
        /// <param name="worksheet">The worksheet to inspect.</param>
        /// <returns>True when a formatting‑only cell is found; otherwise false.</returns>
        public static bool HasFormattingOnlyInitializedCells(Worksheet worksheet)
        {
            // Access the cells collection of the worksheet.
            Cells cells = worksheet.Cells;

            // Determine the used range boundaries.
            int maxRow = cells.MaxRow;
            int maxColumn = cells.MaxColumn;

            // Iterate through every cell in the used range.
            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxColumn; col++)
                {
                    Cell cell = cells[row, col];

                    // Check if the cell has a style explicitly set.
                    if (cell.IsStyleSet)
                    {
                        // Determine if the cell holds no data.
                        bool hasNoValue = cell.Type == CellValueType.IsNull ||
                                          string.IsNullOrEmpty(cell.StringValue);

                        // Exclude cells that contain a formula (even if the result is blank).
                        if (hasNoValue && !cell.IsFormula)
                        {
                            // A formatting‑only cell is found.
                            return true;
                        }
                    }
                }
            }

            // No such cells were detected.
            return false;
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file can be passed as a command‑line argument; otherwise use a default name.
            string filePath = args.Length > 0 ? args[0] : "sample.xlsx";

            // Prevent FileNotFoundException by checking existence first.
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"File not found: {filePath}");
                return;
            }

            try
            {
                // Load the workbook from the specified file.
                Workbook workbook = new Workbook(filePath);

                // Examine each worksheet for formatting‑only cells.
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool hasFormattingOnly = WorksheetUtilities.HasFormattingOnlyInitializedCells(sheet);
                    Console.WriteLine($"Worksheet '{sheet.Name}': Formatting‑only cells present? {hasFormattingOnly}");
                }
            }
            catch (Exception ex)
            {
                // Catch any runtime errors (e.g., corrupted file, unsupported format).
                Console.WriteLine($"Error processing file: {ex.Message}");
            }
        }
    }
}
