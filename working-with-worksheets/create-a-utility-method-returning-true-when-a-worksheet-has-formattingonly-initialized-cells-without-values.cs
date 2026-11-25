// Title: How to detect cells that contain only formatting (no values) in an Aspose.Cells worksheet using C#
// Common Searches: Aspose.Cells C# check for cells with formatting but no data | find empty cells with custom style in Excel using Aspose.Cells | C# method to identify formatting‑only cells in a workbook worksheet | detect cells that have style applied but no value with Aspose.Cells API | scan worksheet for non‑default styled empty cells in C#
// Tags: detect formatting‑only cells Aspose.Cells C# | compare cell style to default workbook style Aspose | scan worksheet for empty styled cells C# | identify non‑default cell formatting without values | worksheet style inspection Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsFormattingOnlyCheck
{
    // Provides a C# utility that iterates through a worksheet up to the last used row/column plus a safety buffer, compares each cell's style with the workbook's default style, and returns true when it finds a cell that has no value (null or empty string) but a non‑default formatting; otherwise it returns false.
    class Program
    {
        /// <param name="worksheet">The worksheet to inspect.</param>
        /// <returns>True if a formatting‑only cell is found; otherwise, false.</returns>
        public static bool HasFormattingOnlyCells(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            try
            {
                Cells cells = worksheet.Cells;

                // Determine a safe scanning range.
                int maxRow = cells.MaxDataRow < 0 ? 0 : cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn < 0 ? 0 : cells.MaxDataColumn;

                const int buffer = 100;
                int scanMaxRow = Math.Min(maxRow + buffer, 1048575); // zero‑based max rows
                int scanMaxCol = Math.Min(maxCol + buffer, 16383);   // zero‑based max columns

                // Default style to compare against.
                Style defaultStyle = worksheet.Workbook.DefaultStyle;

                for (int row = 0; row <= scanMaxRow; row++)
                {
                    for (int col = 0; col <= scanMaxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Check for empty value.
                        bool hasNoValue = cell.Value == null ||
                                          (cell.Value is string s && string.IsNullOrEmpty(s));

                        // Determine if the cell's style differs from the default style.
                        Style style = cell.GetStyle();
                        bool hasFormatting = !style.Equals(defaultStyle);

                        if (hasNoValue && hasFormatting)
                        {
                            // Formatting‑only cell found.
                            return true;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while inspecting worksheet \"{worksheet.Name}\": {ex.Message}");
            }

            // No formatting‑only cells detected.
            return false;
        }

        static void Main(string[] args)
        {
            // Expect the first argument to be the path of the workbook to analyze.
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: AsposeCellsFormattingOnlyCheck <workbook-path>");
                return;
            }

            string workbookPath = args[0];

            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: File not found - {workbookPath}");
                return;
            }

            try
            {
                // Load the workbook.
                Workbook workbook = new Workbook(workbookPath);

                Console.WriteLine($"Analyzing workbook: {workbookPath}");
                Console.WriteLine($"Total worksheets: {workbook.Worksheets.Count}");

                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Worksheet ws = workbook.Worksheets[i];
                    bool hasFormattingOnly = HasFormattingOnlyCells(ws);
                    Console.WriteLine($"Worksheet \"{ws.Name}\": Formatting‑only cells present? {hasFormattingOnly}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
                Console.WriteLine(ex.StackTrace);
            }
        }
    }
}
