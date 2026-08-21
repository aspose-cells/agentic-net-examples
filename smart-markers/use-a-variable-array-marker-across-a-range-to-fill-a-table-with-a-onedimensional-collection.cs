// Title: C# – Generate a One‑Variable Data Table from a One‑Dimensional Array with Aspose.Cells SetTableFormula
// Description: The sample builds a new workbook, places a constant multiplier in B1, writes a formula in C1 that multiplies A1 by that constant, feeds a one‑dimensional array of values, and calls SetTableFormula on C2 to create a single‑row data table. After recalculating, the file is saved as OneDimensionalDataTable.xlsx.
// Keywords: Aspose.Cells C# SetTableFormula | one dimensional data table | variable array marker | Excel automation .NET | C# generate Excel table | sensitivity analysis Excel | global Excel library | USA .NET developers | Europe C# Excel
// Common Searches: SetTableFormula one dimensional collection C# | Aspose.Cells create data table from array | C# generate single‑row data table Excel | How to use variable array marker in Aspose.Cells | One‑variable data table example Aspose.Cells
// Developer Intent: Create a data table that evaluates a formula for each value supplied by a one‑dimensional collection.
// Use Cases: Perform a sensitivity analysis by varying a single input cell across multiple scenarios. | Build a pricing matrix where a base quantity is multiplied by different rates without manual entry. | Automate a scenario‑driven lookup table for financial models that requires rapid recomputation.
// AI Prompts: Show how to switch the table orientation to column‑wise (isRowInput = false) while reusing the same array. | Demonstrate how to read the generated table values back into a C# list after SetTableFormula execution. | Provide a version that accepts a List<double> instead of object[][] for the input collection.

using System;
using Aspose.Cells;

// The sample builds a new workbook, places a constant multiplier in B1, writes a formula in C1 that multiplies A1 by that constant, feeds a one‑dimensional array of values, and calls SetTableFormula on C2 to create a single‑row data table. After recalculating, the file is saved as OneDimensionalDataTable.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put a constant multiplier in B1
        cells["B1"].PutValue(2);

        // Formula that uses the variable input cell A1 and the constant in B1
        // The result will be placed in C1
        cells["C1"].Formula = "=A1*$B$1";

        // One‑dimensional collection of input values for the data table
        // Row count = 1, column count = 5 (values will be placed horizontally)
        object[][] inputValues = new object[1][];
        inputValues[0] = new object[] { 1, 2, 3, 4, 5 };

        // Target cell where the data table will start (below the formula cell)
        Cell targetCell = worksheet.Cells["C2"];

        // Create a one‑variable data table:
        // - 1 row, 5 columns
        // - input cell is A1
        // - input is a row input (isRowInput = true)
        // - values array supplies the input values
        targetCell.SetTableFormula(
            rowNumber: 1,
            columnNumber: 5,
            inputCell: "A1",
            isRowInput: true,
            values: inputValues);

        // Recalculate all formulas so the table shows results
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("OneDimensionalDataTable.xlsx");
    }
}
