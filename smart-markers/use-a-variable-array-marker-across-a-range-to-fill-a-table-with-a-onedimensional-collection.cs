// Title: Create a vertical one‑dimensional data table using SetTableFormula and a variable array smart marker in Aspose.Cells for C#
// AI Prompts: Generate C# code that builds a one‑dimensional double array and uses Worksheet.Cells["C3"].SetTableFormula to fill a column with formula results. | Show how to convert a double[] into the object[][] structure required by SetTableFormula and then recalculate the workbook. | Provide a complete example that defines a formula in B2, applies a variable‑array smart marker, and saves the Excel file.
// Common Searches: Aspose.Cells C# SetTableFormula create data table from array values | How to use a variable array smart marker to populate an Excel column in .NET | One‑variable data table generation with input values from a double array using Aspose.Cells | Convert double array to object[][] for SetTableFormula in C# | Recalculate formulas after applying SetTableFormula in Aspose.Cells
// Tags: Aspose.Cells SetTableFormula vertical data table | C# smart marker array input | one‑dimensional array to object[][] conversion Aspose.Cells | populate Excel column with formula results Aspose.Cells | calculate workbook formulas Aspose.Cells C#

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, sets a formula in B2 that multiplies A1 by 10, builds a one‑dimensional double array, converts it to the object[][] format required by SetTableFormula, applies SetTableFormula starting at C3 to generate a vertical data table, recalculates formulas, prints the results, and saves the file as VariableArrayMarkerOneDimensionalTable.xlsx.
    public class VariableArrayMarkerOneDimensionalTable
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // ------------------------------------------------------------
                // 1. Prepare the source formula.
                //    The formula uses a single input cell (A1) whose value will
                //    be varied by the data table.
                // ------------------------------------------------------------
                // Place the base formula in cell B2 (row index 1, column index 1)
                // Formula: multiply the input cell by 10
                Cell formulaCell = cells[1, 1]; // B2
                formulaCell.Formula = "=A1*10";

                // ------------------------------------------------------------
                // 2. Prepare a one‑dimensional collection of input values.
                //    These values will be fed to the input cell (A1) via the
                //    data table mechanism.
                // ------------------------------------------------------------
                double[] inputValues = new double[] { 1, 2, 3, 4, 5 };

                // Convert the one‑dimensional array to the object[][] format
                // required by SetTableFormula. Each inner array represents a row.
                object[][] tableValues = new object[inputValues.Length][];
                for (int i = 0; i < inputValues.Length; i++)
                {
                    tableValues[i] = new object[] { inputValues[i] };
                }

                // ------------------------------------------------------------
                // 3. Define the target cell where the data table will start.
                //    The table will expand vertically (one column) because we
                //    are using a one‑variable data table (isRowInput = true).
                // ------------------------------------------------------------
                Cell targetCell = worksheet.Cells["C3"]; // Starting point of the table

                // ------------------------------------------------------------
                // 4. Apply the one‑variable data table using SetTableFormula.
                //    Parameters:
                //      rowNumber          = number of rows to populate (inputValues.Length)
                //      columnNumber       = 1 (single column result)
                //      inputCell          = "A1" (the cell whose value changes)
                //      isRowInput         = true  (input values are placed in rows)
                //      values             = prepared object[][]
                // ------------------------------------------------------------
                targetCell.SetTableFormula(
                    rowNumber: inputValues.Length,
                    columnNumber: 1,
                    inputCell: "A1",
                    isRowInput: true,
                    values: tableValues);

                // ------------------------------------------------------------
                // 5. Calculate all formulas so that the table results are materialized.
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 6. Output the generated table to the console for verification.
                //    The table occupies cells C3:C7 (5 rows, 1 column).
                // ------------------------------------------------------------
                Console.WriteLine("One‑dimensional data table results (C3:C7):");
                for (int row = 2; row < 2 + inputValues.Length; row++) // zero‑based index
                {
                    Console.WriteLine($"C{row + 1} = {cells[row, 2].Value}");
                }

                // ------------------------------------------------------------
                // 7. Save the workbook.
                // ------------------------------------------------------------
                workbook.Save("VariableArrayMarkerOneDimensionalTable.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            VariableArrayMarkerOneDimensionalTable.Run();
        }
    }
}
