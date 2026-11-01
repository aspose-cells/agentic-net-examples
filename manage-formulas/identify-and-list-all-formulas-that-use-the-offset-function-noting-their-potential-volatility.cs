// Title: Find and list all Excel formulas that use the OFFSET function and flag them as volatile with Aspose.Cells for .NET (C#)
// Common Searches: C# code to locate volatile OFFSET formulas in an .xlsx using Aspose.Cells | how to enumerate cells containing OFFSET function with Aspose.Cells .NET | retrieve formulas that include OFFSET and mark them volatile in a workbook via Aspose.Cells | list all OFFSET formulas in each worksheet of an Excel file using Aspose.Cells C# | detect volatile functions like OFFSET in Excel with Aspose.Cells library
// Tags: enumerate OFFSET formulas Aspose.Cells | identify volatile Excel functions C# | iterate worksheet used range Aspose.Cells | list formulas with OFFSET Aspose.Cells | flag volatile formulas in .xlsx C#

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook with Aspose.Cells, walks through every worksheet's used range, checks each formula cell for the OFFSET function (case‑insensitive), records the sheet name, cell address, and formula with a '(Volatile)' label, and outputs the collected list of OFFSET formulas.
    class Program
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";

                // Ensure the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"File not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Collection to hold formulas that use OFFSET
                List<string> offsetFormulas = new List<string>();

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Determine the used range of the worksheet
                    AsposeRange usedRange = sheet.Cells.MaxDisplayRange;

                    // Skip empty sheets
                    if (usedRange == null)
                        continue;

                    int startRow = usedRange.FirstRow;
                    int endRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int startCol = usedRange.FirstColumn;
                    int endCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    // Loop through all cells in the used range
                    for (int row = startRow; row <= endRow; row++)
                    {
                        for (int col = startCol; col <= endCol; col++)
                        {
                            Cell cell = sheet.Cells[row, col];

                            // Process only cells that contain a formula
                            if (cell.IsFormula)
                            {
                                string formula = cell.Formula;

                                // Check for the OFFSET function (case‑insensitive)
                                if (formula.IndexOf("OFFSET", StringComparison.OrdinalIgnoreCase) >= 0)
                                {
                                    // OFFSET is a volatile function, so we note its volatility
                                    string entry = $"{sheet.Name}!{cell.Name}: {formula} (Volatile)";
                                    offsetFormulas.Add(entry);
                                }
                            }
                        }
                    }
                }

                // Output the identified formulas
                Console.WriteLine("Formulas that use the OFFSET function:");
                foreach (string item in offsetFormulas)
                {
                    Console.WriteLine(item);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
