// Title: Export Excel formulas with cell addresses to a CSV file using Aspose.Cells for .NET
// Common Searches: C# code to list all formulas in an Excel file and save them as CSV with Aspose.Cells | how to extract cell address and formula from each worksheet using Aspose.Cells .NET | write Excel formula export utility to CSV in C# with Aspose.Cells | retrieve formulas from used range of a workbook with Aspose.Cells for .NET | generate CSV of formula cells from .xlsx using Aspose.Cells library
// Tags: export formulas to CSV Aspose.Cells | iterate used range Aspose.Cells | detect IsFormula property C# | format A1 cell address with sheet name | escape quotes for CSV output C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example loads an .xlsx workbook with Aspose.Cells, walks through each worksheet's MaxDisplayRange, checks the IsFormula flag for each cell, builds a sheet‑prefixed A1 address, escapes double quotes in the formula, and writes "CellAddress,Formula" rows to a CSV file.
class ExportFormulasToCsv
{
    static void Main()
    {
        try
        {
            // Path to the input workbook
            string workbookPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Error: The file '{workbookPath}' was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Path for the CSV output
            string csvPath = "formulas.csv";

            // Ensure the directory for the CSV exists
            string? csvDir = Path.GetDirectoryName(csvPath);
            if (!string.IsNullOrEmpty(csvDir) && !Directory.Exists(csvDir))
            {
                Directory.CreateDirectory(csvDir);
            }

            // Write formulas to CSV
            using (StreamWriter writer = new StreamWriter(csvPath))
            {
                // CSV header
                writer.WriteLine("CellAddress,Formula");

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get the used range of the worksheet
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                    // If the worksheet is empty, skip it
                    if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = startRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = startCol + usedRange.ColumnCount - 1;

                    // Loop through each cell in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                // Build the cell address with sheet name (A1 style)
                                string address = $"'{sheet.Name}'!{cell.Name}";

                                // Escape double quotes for CSV compliance
                                string formula = cell.Formula.Replace("\"", "\"\"");

                                // Write the address and formula, each enclosed in double quotes
                                writer.WriteLine($"\"{address}\",\"{formula}\"");
                            }
                        }
                    }
                }
            }

            Console.WriteLine($"Formulas exported successfully to '{csvPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
