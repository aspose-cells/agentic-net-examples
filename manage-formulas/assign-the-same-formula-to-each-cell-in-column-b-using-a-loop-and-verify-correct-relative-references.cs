// Title: Assign a relative formula to each cell in column B using a loop with Aspose.Cells for .NET
// Common Searches: how to set a formula in column B for each row using Aspose.Cells C# loop | assign =A1+10 formula to column B cells with relative references in Aspose.Cells | verify calculated values after applying formulas with Aspose.Cells .NET | populate column A with sequential numbers and compute column B using Aspose.Cells | save workbook after programmatically adding formulas using Aspose.Cells C#
// Tags: loop assign formulas Aspose.Cells C# | relative cell reference formula Aspose.Cells | calculate workbook formulas .NET | verify formula results Aspose.Cells | save workbook with formulas Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new workbook, fills column A with values 1‑10, then loops through rows assigning the formula "=A{row}+10" to each cell in column B. After calculating the workbook, it verifies that each B cell equals the corresponding A value plus 10, outputs the verification result, and saves the file as ColumnBFormulaDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate column A with sample data (1 to 10)
        for (int i = 0; i < 10; i++)
        {
            // Cells are 0‑based; row i, column 0 corresponds to A{i+1}
            sheet.Cells[i, 0].PutValue(i + 1);
        }

        // Loop through rows 1..10 (0‑based index 0..9) and assign the same formula to column B
        for (int i = 0; i < 10; i++)
        {
            // The formula uses a relative reference to column A of the same row
            // Example for row 1: "=A1+10"
            string formula = $"=A{i + 1}+10";
            sheet.Cells[i, 1].Formula = formula;
        }

        // Calculate the workbook to evaluate the formulas
        workbook.CalculateFormula();

        // Verify that the formulas produced the expected results
        for (int i = 0; i < 10; i++)
        {
            double aValue = sheet.Cells[i, 0].DoubleValue;          // Value from column A
            double bValue = sheet.Cells[i, 1].DoubleValue;          // Calculated value in column B
            double expected = aValue + 10;                         // Expected result

            if (Math.Abs(bValue - expected) > 1e-9)
            {
                Console.WriteLine($"Row {i + 1}: Verification failed. Expected {expected}, got {bValue}");
            }
            else
            {
                Console.WriteLine($"Row {i + 1}: Verification succeeded. Value = {bValue}");
            }
        }

        // Save the workbook (optional, demonstrates saving)
        workbook.Save("ColumnBFormulaDemo.xlsx");
    }
}
