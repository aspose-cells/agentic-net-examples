// Title: C# – Convert Excel Workbook to CSV with Formulas Preserved as Text using Aspose.Cells
// Description: Load an Excel workbook with Aspose.Cells, replace every formula cell with its literal formula string via PutValue, and save the workbook as a CSV file. The resulting CSV keeps formulas intact as text, ideal for audits, data migration, or downstream processing.
// Keywords: Aspose.Cells | C# | Excel to CSV conversion | preserve formulas as text | formula string export | SaveFormat.Csv | Workbook.Save CSV | .NET Excel automation | cell.IsFormula | PutValue | data migration | financial model audit | global
// Common Searches: Aspose.Cells export Excel to CSV with formulas as text | C# convert workbook to CSV keeping formula strings | how to save Excel formulas in CSV using Aspose.Cells | replace formula cells with formula text before CSV export | batch convert multiple worksheets to CSV preserving formulas
// Developer Intent: Generate a CSV file from an Excel workbook where each formula cell is written as its original formula text rather than the evaluated value.
// Use Cases: Auditing financial spreadsheets by exporting formulas to a readable CSV format. | Migrating Excel data to a plain‑text system while retaining the original calculation logic. | Feeding CSV files to downstream applications that require the formula syntax for re‑creation of the workbook.
// AI Prompts: Create C# code that loads an Excel file with Aspose.Cells, converts all formula cells to their formula strings, and saves the result as CSV. | Show how to modify the conversion to target only selected worksheets or a specific cell range while preserving formulas as text. | Explain which SaveOptions to set for CSV export (encoding, delimiter) when exporting formulas using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Load an Excel workbook with Aspose.Cells, replace every formula cell with its literal formula string via PutValue, and save the workbook as a CSV file. The resulting CSV keeps formulas intact as text, ideal for audits, data migration, or downstream processing.
    public class WorkbookToCsvWithFormulas
    {
        /// <param name="sourcePath">Path to the source Excel file.</param>
        /// <param name="csvPath">Path where the resulting CSV file will be saved.</param>
        public static void Convert(string sourcePath, string csvPath)
        {
            try
            {
                // Load the workbook from the specified file (lifecycle: load)
                Workbook workbook = new Workbook(sourcePath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the maximum used row and column indices
                    int maxRow = sheet.Cells.MaxDataRow;
                    int maxCol = sheet.Cells.MaxDataColumn;

                    // Loop over all used cells
                    for (int row = 0; row <= maxRow; row++)
                    {
                        for (int col = 0; col <= maxCol; col++)
                        {
                            var cell = sheet.Cells[row, col];

                            // If the cell contains a formula, replace it with the formula text
                            if (cell.IsFormula)
                            {
                                // Preserve the formula as a string value
                                string formulaText = cell.Formula;
                                cell.PutValue(formulaText);
                            }
                        }
                    }
                }

                // Save the modified workbook as CSV (lifecycle: save)
                workbook.Save(csvPath, SaveFormat.Csv);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during conversion: {ex.Message}");
                throw;
            }
        }

        // Example usage
        public static void Run(string sourceFile, string outputCsv)
        {
            if (!File.Exists(sourceFile))
            {
                Console.WriteLine($"Source file not found: {sourceFile}");
                return;
            }

            Convert(sourceFile, outputCsv);
            Console.WriteLine($"Workbook converted to CSV with formulas preserved at: {outputCsv}");
        }

        // Entry point
        public static void Main(string[] args)
        {
            try
            {
                string sourceFile = args.Length > 0 ? args[0] : "input.xlsx";
                string outputCsv = args.Length > 1 ? args[1] : "output.csv";

                Run(sourceFile, outputCsv);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}
