using System;
using System.IO;
using System.Text.RegularExpressions;
using Aspose.Cells;

namespace AsposeCellsFormulaCopy
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate source range A1:C3 with sample data and formulas
                // A1 = 1, B1 = 2, C1 = 3
                // A2 = =A1+1 (relative), B2 = =B1+1, C2 = =C1+1
                // A3 = =A2+1, etc.
                for (int row = 0; row < 3; row++)
                {
                    for (int col = 0; col < 3; col++)
                    {
                        if (row == 0)
                        {
                            // First row: put numeric values
                            cells[row, col].PutValue(row * 3 + col + 1);
                        }
                        else
                        {
                            // Subsequent rows: set a simple relative formula referencing the cell above
                            string colLetter = CellsHelper.ColumnIndexToName(col);
                            string formula = $"={colLetter}{row}+1"; // e.g., =A1+1
                            cells[row, col].Formula = formula;
                        }
                    }
                }

                // Define source and destination ranges (use fully qualified Aspose.Cells.Range to avoid ambiguity)
                Aspose.Cells.Range sourceRange = cells.CreateRange("A1:C3");
                Aspose.Cells.Range destRange = cells.CreateRange("E1:G3");

                // Iterate through each cell in the source range,
                // convert its formula to absolute references, and set it in the destination range.
                for (int i = 0; i < sourceRange.RowCount; i++)
                {
                    for (int j = 0; j < sourceRange.ColumnCount; j++)
                    {
                        Cell srcCell = sourceRange[i, j];
                        Cell dstCell = destRange[i, j];

                        if (srcCell.IsFormula)
                        {
                            // Get the original formula (A1 style)
                            string originalFormula = srcCell.Formula;

                            // Convert relative references (e.g., A1) to absolute references (e.g., $A$1)
                            // Regex adds $ before column letters and row numbers.
                            string absoluteFormula = Regex.Replace(
                                originalFormula,
                                @"([A-Z]+)(\d+)",
                                @"$$$1$$$2");

                            // Set the absolute formula in the destination cell.
                            dstCell.SetFormula(absoluteFormula, new FormulaParseOptions());
                        }
                        else
                        {
                            // If the source cell does not contain a formula, copy its value.
                            dstCell.PutValue(srcCell.Value);
                        }
                    }
                }

                // Recalculate formulas to reflect the new values.
                workbook.CalculateFormula();

                // Save the workbook (ensure the directory exists)
                string outputPath = "FormulaCopyAbsolute.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}