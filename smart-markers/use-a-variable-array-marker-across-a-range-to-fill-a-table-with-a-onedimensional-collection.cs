// Title: Aspose.Cells C# – Create a One‑Dimensional Data Table with Variable Array Marker (SetTableFormula)
// Description: Demonstrates how to generate an Excel data table from a one‑dimensional collection using Aspose.Cells' variable array marker. The example builds a workbook, defines a base formula, converts a double[] to the object[][] format required by SetTableFormula, applies the formula with row‑input mode, recalculates, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | SetTableFormula | variable array marker | one dimensional array | Excel data table | row input variable | financial modeling | sensitivity analysis | spreadsheet automation
// Common Searches: Aspose.Cells SetTableFormula one dimensional array | variable array marker C# example | populate Excel column from double array Aspose.Cells | create data table with row input variable Aspose.Cells | how to use SetTableFormula in .NET
// Developer Intent: Generate an Excel table by applying a formula to each element of a one‑dimensional collection using a variable array marker.
// Use Cases: Run a sensitivity analysis where a list of rates drives a calculation across rows. | Build a lookup table for financial scenarios without writing individual formulas. | Automate the creation of a column of results from a collection of input parameters.
// AI Prompts: Write C# code that uses Aspose.Cells SetTableFormula to create a data table from a string[] with a column input variable. | Explain how to convert any one‑dimensional collection into the object[][] structure required by SetTableFormula. | Show an example of using SetTableFormula with isRowInput set to false to fill a table horizontally.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to generate an Excel data table from a one‑dimensional collection using Aspose.Cells' variable array marker. The example builds a workbook, defines a base formula, converts a double[] to the object[][] format required by SetTableFormula, applies the formula with row‑input mode, recalculates, and saves the file.
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

                // Set a base formula that will be used for the data table.
                // Example: multiply the input value (A1) by 10.
                cells["B1"].Formula = "=A1*10";

                // Prepare a one‑dimensional collection (e.g., an array of interest rates)
                double[] inputValues = new double[] { 0.05, 0.10, 0.15, 0.20 };

                // Convert the one‑dimensional collection to the object[][] format required by SetTableFormula.
                // Each inner array represents a row in the table; here we have a single column.
                object[][] tableValues = new object[inputValues.Length][];
                for (int i = 0; i < inputValues.Length; i++)
                {
                    tableValues[i] = new object[] { inputValues[i] };
                }

                // Choose the cell where the data table will start (top‑left corner of the result range).
                // The table will spill downwards because we are using a row input variable.
                Cell tableStartCell = cells["B2"];

                // Create a one‑variable data table:
                // - rowNumber: number of rows to populate (same as the number of input values)
                // - columnNumber: 1 column because we have a single result column
                // - inputCell: the cell that serves as the input variable for the formula ("A1")
                // - isRowInput: true indicates that the input cell is a row input (values vary down the rows)
                // - values: the prepared one‑dimensional collection
                tableStartCell.SetTableFormula(
                    rowNumber: inputValues.Length,
                    columnNumber: 1,
                    inputCell: "A1",
                    isRowInput: true,
                    values: tableValues
                );

                // Recalculate all formulas so that the table results are materialized.
                workbook.CalculateFormula();

                // Optional: display the generated table values in the console for verification.
                Console.WriteLine("Generated one‑dimensional data table (B2:B5):");
                for (int i = 0; i < inputValues.Length; i++)
                {
                    Console.WriteLine($"B{2 + i} = {cells[1 + i, 1].Value}");
                }

                // Save the workbook to a file.
                string outputPath = "VariableArrayMarkerOneDimensionalTable.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
