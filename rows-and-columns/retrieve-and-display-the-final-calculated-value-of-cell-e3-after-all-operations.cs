// Title: C# – Get the calculated value of cell E3 after formula evaluation with Aspose.Cells
// Description: Creates a workbook, sets numeric values in A1, B1, C1, defines formulas in D3 (=A1+B1*C1) and E3 (=D3*2), runs Workbook.CalculateFormula(), then reads worksheet.Cells["E3"].Value to output the final result (50).
// Keywords: Aspose.Cells | C# | calculate formula | read cell value | formula result | E3 | Workbook.CalculateFormula | dependent formulas | retrieve calculated value | Aspose.Cells .NET example
// Common Searches: Aspose.Cells get formula result | How to read calculated cell value in C# | Retrieve value of cell after CalculateFormula | Aspose.Cells dependent formula value | C# example reading E3 after formulas
// Developer Intent: Show how to obtain the evaluated result of a dependent formula cell (E3) using Aspose.Cells for .NET.
// Use Cases: Display the computed value in a console or UI component. | Pass the result to further calculations or business logic. | Store the value in a database or another worksheet for reporting. | Use the value in conditional logic or API responses. | Format the result for presentation (e.g., currency or percentage).
// AI Prompts: Write C# code that calculates formulas and returns the value of a specific cell using Aspose.Cells. | Show how to format the retrieved E3 value as currency or percentage. | Explain error handling when a formula cell returns #DIV/0! or null after calculation. | Demonstrate how to batch read multiple calculated cells efficiently with Aspose.Cells. | Provide a unit test that verifies the E3 value equals 50.

using System;
using Aspose.Cells;

// Creates a workbook, sets numeric values in A1, B1, C1, defines formulas in D3 (=A1+B1*C1) and E3 (=D3*2), runs Workbook.CalculateFormula(), then reads worksheet.Cells["E3"].Value to output the final result (50).
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some cells that will be used in formulas
        worksheet.Cells["A1"].PutValue(5);
        worksheet.Cells["B1"].PutValue(10);
        worksheet.Cells["C1"].PutValue(2);

        // Set a formula in D3 that depends on A1, B1, and C1
        worksheet.Cells["D3"].Formula = "=A1+B1*C1"; // Expected result: 5 + 10*2 = 25

        // Set a formula in E3 that depends on D3
        worksheet.Cells["E3"].Formula = "=D3*2"; // Expected result: 25 * 2 = 50

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve and display the final calculated value of cell E3
        object e3Value = worksheet.Cells["E3"].Value;
        Console.WriteLine("Final calculated value of E3: " + e3Value);
    }
}
