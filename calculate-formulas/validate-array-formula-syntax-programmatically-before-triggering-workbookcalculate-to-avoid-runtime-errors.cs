using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    class ValidateArrayFormulas
    {
        static void Main()
        {
            try
            {
                // ---------- Create a new workbook ----------
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // ---------- Populate some data ----------
                cells["A1"].PutValue(1);
                cells["A2"].PutValue(2);
                cells["A3"].PutValue(3);

                // ---------- Set a valid array formula ----------
                // This will fill B1:B3 with the transposed values of A1:A3
                cells["B1"].SetArrayFormula("=TRANSPOSE(A1:A3)", 3, 1);

                // ---------- Set an invalid array formula (syntax error) ----------
                // Wrap in try‑catch to prevent exception from bubbling up
                try
                {
                    cells["C1"].SetArrayFormula("=SUM(A1:A3", 1, 1);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set array formula in C1: {ex.Message}");
                    // Leave the cell empty so later validation can handle it
                    cells["C1"].Formula = string.Empty;
                }

                // ---------- Validate all array formulas before full calculation ----------
                foreach (Cell cell in worksheet.Cells)
                {
                    if (cell.IsArrayFormula)
                    {
                        string formula = cell.Formula;
                        try
                        {
                            // Attempt to calculate the array formula in isolation.
                            // If the formula is syntactically correct, this will succeed.
                            object[][] dummyResult = worksheet.CalculateArrayFormula(formula, new CalculationOptions());

                            // Optional: inspect dummyResult here if needed.
                        }
                        catch (Exception ex)
                        {
                            // The formula is invalid – report and clear it to avoid runtime errors.
                            Console.WriteLine($"Invalid array formula detected in cell {cell.Name}: {formula}");
                            Console.WriteLine($"Error details: {ex.Message}");

                            // Clear the faulty formula so that Workbook.CalculateFormula won't fail.
                            cell.Formula = string.Empty;
                        }
                    }
                }

                // ---------- Perform normal workbook calculation (now safe) ----------
                workbook.CalculateFormula();

                // ---------- Save the workbook ----------
                string outputPath = "ValidatedArrayFormulas.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }
    }
}