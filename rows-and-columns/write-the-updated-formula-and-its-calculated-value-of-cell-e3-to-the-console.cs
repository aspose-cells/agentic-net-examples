// Title: Print Updated Formula and Calculated Value of Cell E3 to Console with Aspose.Cells for .NET
// Description: Creates a workbook, fills A1, A2, B1, and B2 with numbers, assigns the formula =SUM(A1:A2)+SUM(B1:B2) to E3, recalculates the sheet, and writes both the formula string and the evaluated result of E3 to the console.
// Keywords: Aspose.Cells C# formula | retrieve cell formula .NET | calculate workbook formulas | print cell value console | E3 formula Aspose.Cells | SUM function example | Aspose.Cells CalculateFormula
// Common Searches: Aspose.Cells get cell formula after calculation | C# Aspose.Cells print cell value | display formula and result in console using Aspose.Cells | Aspose.Cells CalculateFormula example | retrieve cells["E3"].Value .NET
// Developer Intent: Show how to read a cell’s formula and its computed value after calling Workbook.CalculateFormula() and output both to the console.
// Use Cases: Debugging: log the exact formula and result of a critical cell after workbook calculation. | Reporting: output financial or statistical totals directly from a console app without opening the workbook. | Verification: create a quick sanity‑check that confirms formulas evaluate as expected during automated tests.
// AI Prompts: Generate C# code that sets a SUM formula in cell E3 with Aspose.Cells, runs Workbook.CalculateFormula(), and prints the formula and its value to the console. | Provide an Aspose.Cells example that populates A1‑B2, assigns =SUM(A1:A2)+SUM(B1:B2) to E3, recalculates, and displays both cells["E3"].Formula and cells["E3"].Value. | Write a snippet demonstrating how to retrieve and log cells["E3"].Formula and cells["E3"].Value after invoking CalculateFormula in a .NET console application.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    // Creates a workbook, fills A1, A2, B1, and B2 with numbers, assigns the formula =SUM(A1:A2)+SUM(B1:B2) to E3, recalculates the sheet, and writes both the formula string and the evaluated result of E3 to the console.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some cells that will be referenced by the formula
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(5);
            cells["B2"].PutValue(15);

            // Set a formula in cell E3 that uses the above values
            // Example: sum of A1:A2 plus sum of B1:B2
            cells["E3"].Formula = "=SUM(A1:A2) + SUM(B1:B2)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the updated formula and its calculated value
            Console.WriteLine("Updated formula in E3: " + cells["E3"].Formula);
            Console.WriteLine("Calculated value in E3: " + cells["E3"].Value);
        }
    }
}
