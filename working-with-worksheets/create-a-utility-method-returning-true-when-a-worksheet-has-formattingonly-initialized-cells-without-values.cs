using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUtilities
{
    public static class WorksheetHelper
    {
        /// <summary>
        /// Determines whether the specified worksheet contains any cells that have a style applied
        /// but do not contain a value, formula, or any other data.
        /// </summary>
        /// <param name="worksheet">The worksheet to inspect.</param>
        /// <returns>True if at least one formatting‑only cell is found; otherwise, false.</returns>
        public static bool HasFormattingOnlyCells(Worksheet worksheet)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));

            Cells cells = worksheet.Cells;

            // Determine the bounds to scan. MaxDataRow/MaxDataColumn cover cells that have data.
            // Formatting‑only cells may exist outside this range, so also include the last used row/column.
            int maxRow = Math.Max(cells.MaxDataRow, cells.MaxRow);
            int maxCol = Math.Max(cells.MaxDataColumn, cells.MaxColumn);

            for (int row = 0; row <= maxRow; row++)
            {
                for (int col = 0; col <= maxCol; col++)
                {
                    Cell cell = cells[row, col];

                    // Cell has a style explicitly set (not default)
                    bool styleSet = cell.IsStyleSet;

                    // Cell contains a value (including empty string) or a formula
                    bool hasValue = cell.Value != null && !(cell.Value is string s && string.IsNullOrEmpty(s));
                    bool hasFormula = cell.IsFormula;

                    // Formatting‑only means style is set but no value/formula is present
                    if (styleSet && !hasValue && !hasFormula)
                    {
                        return true;
                    }
                }
            }

            return false;
        }
    }

    internal class Program
    {
        /// <summary>
        /// Entry point for the console application.
        /// Loads a workbook (if the file exists) and reports formatting‑only cells per worksheet.
        /// </summary>
        /// <param name="args">Optional first argument: path to the Excel file.</param>
        private static void Main(string[] args)
        {
            try
            {
                // Determine workbook path (default to "sample.xlsx" if not provided)
                string workbookPath = args.Length > 0 ? args[0] : "sample.xlsx";

                // Prevent FileNotFoundException
                if (!File.Exists(workbookPath))
                {
                    Console.WriteLine($"File not found: {workbookPath}");
                    return;
                }

                // Load workbook
                Workbook workbook = new Workbook(workbookPath);

                // Inspect each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    bool hasFormattingOnly = WorksheetHelper.HasFormattingOnlyCells(sheet);
                    Console.WriteLine($"Worksheet '{sheet.Name}' has formatting‑only cells: {hasFormattingOnly}");
                }
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}