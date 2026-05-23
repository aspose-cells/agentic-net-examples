using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class FormulaTextExtractor
    {
        static void Main()
        {
            try
            {
                // Input workbook path
                string inputPath = "input.xlsx";

                // Verify that the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);
                Worksheet worksheet = workbook.Worksheets[0];

                // Define the range to extract formulas from (e.g., A1:C10)
                string startCell = "A1";
                string endCell = "C10";
                Aspose.Cells.Range range = worksheet.Cells.CreateRange(startCell, endCell);

                // Iterate through each cell in the range and output its formula text
                for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                {
                    for (int c = range.FirstColumn; c < range.FirstColumn + range.ColumnCount; c++)
                    {
                        Cell cell = worksheet.Cells[r, c];

                        // Process only cells that contain a formula
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            // Get the formula text in A1 notation (non‑R1C1, non‑local)
                            string formulaText = cell.GetFormula(false, false);
                            Console.WriteLine($"{cell.Name}: {formulaText}");
                        }
                    }
                }

                // OPTIONAL: Write the extracted formulas to a new column for persistence
                int outputColumn = range.FirstColumn + range.ColumnCount + 1; // column after the source range
                int outputRow = range.FirstRow;

                for (int r = range.FirstRow; r < range.FirstRow + range.RowCount; r++)
                {
                    for (int c = range.FirstColumn; c < range.FirstColumn + range.ColumnCount; c++)
                    {
                        Cell cell = worksheet.Cells[r, c];
                        if (!string.IsNullOrEmpty(cell.Formula))
                        {
                            string formulaText = cell.GetFormula(false, false);
                            worksheet.Cells[outputRow, outputColumn].PutValue(formulaText);
                            outputRow++;
                        }
                    }
                }

                // Save the workbook with the extracted formulas
                string outputPath = "output.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                // Runtime safety: report any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}