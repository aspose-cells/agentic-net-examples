// Title: C# console app to compute the percentage of initialized cells that contain formulas in an Excel worksheet using Aspose.Cells
// Common Searches: C# program to calculate formula cell ratio in a specific Excel sheet with Aspose.Cells | how to determine formula density for a worksheet using Aspose.Cells .NET | Aspose.Cells count formula cells versus total initialized cells in used range | measure Excel worksheet formula complexity in a C# console application | retrieve initialized cell count and formula count from a workbook with Aspose.Cells
// Tags: formula cell ratio Aspose.Cells | initialized cell count Aspose.Cells | used range traversal Aspose.Cells | worksheet selection by name or index Aspose.Cells | excel formula complexity metric .NET

using System;
using System.IO;
using Aspose.Cells;

namespace FormulaComplexityCalculator
{
    // The example loads an Excel file, selects a worksheet by name or index, iterates over its used range counting all initialized cells and those that contain formulas (via IsFormula or a leading '='), computes the formula‑cell percentage, and prints the worksheet name, total initialized cells, formula cells, and density, with error handling for missing files and invalid worksheets.
    class Program
    {
        /// <param name="args">
        /// args[0] - path to the Excel file (required)
        /// args[1] - worksheet name or zero‑based index (optional, defaults to first worksheet)
        /// </param>
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Usage: FormulaComplexityCalculator <excelFilePath> [worksheetNameOrIndex]");
                return;
            }

            string filePath = args[0];
            if (!File.Exists(filePath))
            {
                Console.WriteLine($"Error: File not found – {filePath}");
                return;
            }

            try
            {
                // Load the workbook
                Workbook workbook = new Workbook(filePath);

                // Resolve the worksheet
                Worksheet worksheet = ResolveWorksheet(workbook, args);
                if (worksheet == null)
                {
                    // ResolveWorksheet already printed an error message
                    return;
                }

                // Get the cells collection
                Cells cells = worksheet.Cells;

                // Determine the used range using MaxDataRow/MaxDataColumn
                int maxRow = cells.MaxDataRow;    // -1 if no data
                int maxCol = cells.MaxDataColumn; // -1 if no data

                if (maxRow < 0 || maxCol < 0)
                {
                    Console.WriteLine("The worksheet contains no initialized cells.");
                    return;
                }

                long totalInitialized = 0;
                long formulaCount = 0;

                // Iterate through the used range
                for (int row = 0; row <= maxRow; row++)
                {
                    for (int col = 0; col <= maxCol; col++)
                    {
                        Cell cell = cells[row, col];

                        // Skip cells that have no value (null)
                        if (cell.Type == CellValueType.IsNull)
                            continue;

                        totalInitialized++;

                        // Determine if the cell contains a formula
                        bool hasFormula = cell.IsFormula ||
                                          (!string.IsNullOrEmpty(cell.Formula) && cell.Formula[0] == '=');

                        if (hasFormula)
                            formulaCount++;
                    }
                }

                if (totalInitialized == 0)
                {
                    Console.WriteLine("No initialized cells were found in the examined range.");
                    return;
                }

                double percentage = (double)formulaCount / totalInitialized * 100.0;
                Console.WriteLine($"Worksheet: {worksheet.Name}");
                Console.WriteLine($"Initialized cells: {totalInitialized}");
                Console.WriteLine($"Cells with formulas: {formulaCount}");
                Console.WriteLine($"Formula density: {percentage:F2}%");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while processing the workbook: {ex.Message}");
            }
        }

        private static Worksheet ResolveWorksheet(Workbook workbook, string[] args)
        {
            try
            {
                if (args.Length >= 2)
                {
                    string wsIdentifier = args[1];
                    // Try to parse as integer index first
                    if (int.TryParse(wsIdentifier, out int wsIndex))
                    {
                        if (wsIndex < 0 || wsIndex >= workbook.Worksheets.Count)
                        {
                            Console.WriteLine($"Error: Worksheet index {wsIndex} is out of range.");
                            return null;
                        }
                        return workbook.Worksheets[wsIndex];
                    }
                    else
                    {
                        // Treat as name
                        Worksheet ws = workbook.Worksheets[wsIdentifier];
                        if (ws == null)
                        {
                            Console.WriteLine($"Error: Worksheet named \"{wsIdentifier}\" does not exist.");
                            return null;
                        }
                        return ws;
                    }
                }
                else
                {
                    // Default to first worksheet
                    return workbook.Worksheets[0];
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error while resolving worksheet: {ex.Message}");
                return null;
            }
        }
    }
}
