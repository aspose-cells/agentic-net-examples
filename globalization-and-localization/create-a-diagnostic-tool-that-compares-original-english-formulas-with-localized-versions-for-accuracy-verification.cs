using System;
using System.IO;
using Aspose.Cells;
using CellsRange = Aspose.Cells.Range; // Alias to avoid conflict with System.Range

namespace FormulaLocalizationDiagnostic
{
    public class DiagnosticTool
    {
        // Runs the diagnostic: loads a workbook, compares English and localized formulas,
        // writes the comparison to a new worksheet, and saves the result.
        public static void Run(string inputPath, string outputPath)
        {
            try
            {
                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                    return;
                }

                // Load the workbook from the specified file
                Workbook workbook = new Workbook(inputPath);

                // Create a new worksheet to store diagnostic results
                Worksheet diagSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                diagSheet.Name = "FormulaDiagnostic";

                // Write header row
                diagSheet.Cells["A1"].PutValue("Sheet");
                diagSheet.Cells["B1"].PutValue("Cell");
                diagSheet.Cells["C1"].PutValue("English Formula");
                diagSheet.Cells["D1"].PutValue("Localized Formula");
                diagSheet.Cells["E1"].PutValue("Match (English=Localized)");

                int outputRow = 1; // zero‑based index; start after header

                // Iterate through all worksheets and cells
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Skip the diagnostic sheet itself
                    if (ws.Name == diagSheet.Name) continue;

                    // Get the used range to limit iteration
                    CellsRange usedRange = ws.Cells.MaxDisplayRange;
                    if (usedRange == null) continue;

                    int firstRow = usedRange.FirstRow;
                    int lastRow = usedRange.FirstRow + usedRange.RowCount - 1;
                    int firstCol = usedRange.FirstColumn;
                    int lastCol = usedRange.FirstColumn + usedRange.ColumnCount - 1;

                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        for (int col = firstCol; col <= lastCol; col++)
                        {
                            Cell cell = ws.Cells[row, col];

                            // Process only cells that contain a formula
                            if (!string.IsNullOrEmpty(cell.Formula))
                            {
                                string englishFormula = cell.Formula;          // Standard (en‑US) formula
                                string localizedFormula = cell.FormulaLocal;   // Locale‑specific formula

                                bool match = string.Equals(englishFormula, localizedFormula, StringComparison.OrdinalIgnoreCase);

                                // Write the comparison data to the diagnostic sheet
                                diagSheet.Cells[outputRow, 0].PutValue(ws.Name);
                                diagSheet.Cells[outputRow, 1].PutValue(cell.Name);
                                diagSheet.Cells[outputRow, 2].PutValue(englishFormula);
                                diagSheet.Cells[outputRow, 3].PutValue(localizedFormula);
                                diagSheet.Cells[outputRow, 4].PutValue(match ? "Yes" : "No");

                                outputRow++;
                            }
                        }
                    }
                }

                // Auto‑fit columns for readability
                diagSheet.AutoFitColumns();

                // Save the workbook with diagnostics
                workbook.Save(outputPath);
                Console.WriteLine($"Diagnostic completed. Results saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during diagnostic: {ex.Message}");
            }
        }

        // Example usage
        public static void Main(string[] args)
        {
            // Ensure two arguments: input file path and output file path
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: FormulaLocalizationDiagnostic <input.xlsx> <output.xlsx>");
                return;
            }

            string inputFile = args[0];
            string outputFile = args[1];

            Run(inputFile, outputFile);
        }
    }
}