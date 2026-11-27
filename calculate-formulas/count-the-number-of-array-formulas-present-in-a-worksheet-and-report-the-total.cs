// Title: Counting array‑formula cells in a selected worksheet using Aspose.Cells for .NET
// Common Searches: aspocells c# count array formulas in a worksheet | how to get number of array formula cells with Aspose.Cells .NET | enumerate array formula cells in a specific sheet using Aspose.Cells | C# example to report total array formulas in an Excel file with Aspose.Cells | retrieve array formula count from worksheet index using Aspose.Cells
// Tags: array formula counting Aspose.Cells | Cell.IsArrayFormula property | iterate used range cells C# | select worksheet by index Aspose.Cells | total array formulas report .NET

using System;
using Aspose.Cells;

namespace ArrayFormulaCounter
{
    // The sample loads an Excel workbook with Aspose.Cells, optionally selects a worksheet by its zero‑based index, iterates over the used cell range via MaxDataRow/MaxDataColumn, checks each cell's IsArrayFormula flag, increments a counter for matching cells, and outputs the total number of array‑formula cells in the chosen worksheet while handling invalid arguments and runtime errors.
    class Program
    {
        /// <param name="args">
        /// Optional command‑line arguments.
        /// args[0] – path to the input Excel file (required).
        /// args[1] – (optional) zero‑based index of the worksheet to inspect. Default is 0.
        /// </param>
        static void Main(string[] args)
        {
            // Validate input arguments
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: ArrayFormulaCounter <input-file.xlsx> [worksheet-index]");
                return;
            }

            string inputPath = args[0];
            int sheetIndex = 0;

            if (args.Length > 1 && !int.TryParse(args[1], out sheetIndex))
            {
                Console.WriteLine("Invalid worksheet index supplied. Using the first worksheet (index 0).");
                sheetIndex = 0;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Ensure the requested worksheet exists
                if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
                {
                    Console.WriteLine($"Worksheet index {sheetIndex} is out of range. The workbook contains {workbook.Worksheets.Count} worksheet(s).");
                    return;
                }

                Worksheet worksheet = workbook.Worksheets[sheetIndex];
                Cells cells = worksheet.Cells;

                // Iterate through all used cells and count array formulas
                int arrayFormulaCount = 0;

                // Using the MaxDataRow/MaxDataColumn properties limits the loop to the used range,
                // which is far more efficient than scanning the entire sheet.
                int maxRow = cells.MaxDataRow;
                int maxCol = cells.MaxDataColumn;

                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];
                        // The Cell.IsArrayFormula property returns true if the cell belongs to an array formula.
                        if (cell.IsArrayFormula)
                        {
                            arrayFormulaCount++;
                        }
                    }
                }

                Console.WriteLine($"Worksheet \"{worksheet.Name}\" contains {arrayFormulaCount} array formula cell(s).");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing file \"{inputPath}\": {ex.Message}");
            }
        }
    }
}
