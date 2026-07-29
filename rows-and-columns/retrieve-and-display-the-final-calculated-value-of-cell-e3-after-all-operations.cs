// Title: Aspose.Cells .NET – Retrieve the final calculated value of cell E3 after formulas
// Description: Creates a workbook, assigns numeric values to A1 and B1, sets formulas in C1, D1 and E3, runs workbook.CalculateFormula(), then reads and prints the computed value of E3 (130).
// Keywords: Aspose.Cells get calculated cell value | read formula result C# | Workbook.CalculateFormula example | retrieve cell E3 value Aspose | Aspose.Cells .NET formula evaluation
// Common Searches: how to read calculated value of a cell in Aspose.Cells | Aspose.Cells C# get result of formula in E3 | calculate all formulas and fetch cell value Aspose | Aspose.Cells retrieve cell value after workbook.CalculateFormula
// Developer Intent: Obtain the evaluated value of cell E3 after all dependent formulas are calculated.
// Use Cases: Show a total or summary stored in E3 after user inputs change. | Export the final amount from E3 to a reporting or billing system. | Log the computed E3 value for debugging or audit trails.
// AI Prompts: Generate C# code that calculates all formulas in an Aspose.Cells workbook and returns the value of cell E3. | Explain how to ensure formulas are refreshed before reading a cell value with Aspose.Cells .NET. | Provide error‑handling examples when retrieving a calculated cell value from Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Creates a workbook, assigns numeric values to A1 and B1, sets formulas in C1, D1 and E3, runs workbook.CalculateFormula(), then reads and prints the computed value of E3 (130).
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some data that will be used in formulas
            cells["A1"].PutValue(5);               // A1 = 5
            cells["B1"].PutValue(10);              // B1 = 10

            // Set formulas that depend on the above values
            cells["C1"].Formula = "=A1+B1";        // C1 = 15
            cells["D1"].Formula = "=C1*2";         // D1 = 30

            // The target cell E3 whose final value we want to retrieve
            cells["E3"].Formula = "=D1+100";       // E3 = 130

            // Calculate all formulas in the workbook (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // Retrieve the calculated value of cell E3
            object e3Value = cells["E3"].Value;

            // Display the result
            Console.WriteLine("Final calculated value of E3: " + e3Value);
        }
    }
}
