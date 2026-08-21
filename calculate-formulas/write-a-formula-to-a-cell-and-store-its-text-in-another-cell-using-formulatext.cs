// Title: C# – Write a Formula and Retrieve Its Text with FORMULATEXT using Aspose.Cells
// Description: Shows how to assign a formula to cell A1, use the Excel FORMULATEXT function in cell B1 to capture the formula string, calculate all formulas, display the results, and save the workbook as FormulaTextDemo.xlsx with Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# | FORMULATEXT | write formula | retrieve formula text | calculate formulas | Excel formula as string | save workbook | Aspose.Cells .NET example
// Common Searches: Aspose.Cells FORMULATEXT example C# | How to get formula string from a cell using Aspose.Cells | Write formula to cell and display formula text Aspose | Calculate formulas and export workbook Aspose.Cells | Retrieve Excel formula text programmatically .NET
// Developer Intent: Create a workbook, set a formula in one cell, display that formula as text in another cell with FORMULATEXT, evaluate the workbook, and save the file.
// Use Cases: Show the original formula next to its calculated result for auditing. | Generate documentation that lists each formula used in a spreadsheet. | Export a report where formulas are displayed as plain text rather than evaluated values. | Validate that formulas were entered correctly by comparing the stored text with expected expressions.
// AI Prompts: Provide a C# Aspose.Cells snippet that writes "=SUM(10,20,30)" to A1, uses FORMULATEXT to copy the formula text to B1, calculates the workbook, prints both values, and saves the file. | Explain how to retrieve the formula text from a cell with Aspose.Cells and handle cases where the referenced cell contains no formula. | Show how to loop through a range, capture each cell's formula using FORMULATEXT, and write the collected texts to a summary worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsFormulaTextDemo
{
    // Shows how to assign a formula to cell A1, use the Excel FORMULATEXT function in cell B1 to capture the formula string, calculate all formulas, display the results, and save the workbook as FormulaTextDemo.xlsx with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Write a formula to cell A1
            cells["A1"].Formula = "=SUM(10,20,30)";

            // Write a formula to cell B1 that returns the text of the formula in A1
            // FORMULATEXT is an Excel function that returns the formula as a string
            cells["B1"].Formula = "=FORMULATEXT(A1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the results
            Console.WriteLine("A1 calculated value: " + cells["A1"].StringValue); // Should be 60
            Console.WriteLine("B1 formula text: " + cells["B1"].StringValue);   // Should be =SUM(10,20,30)

            // Save the workbook (optional)
            workbook.Save("FormulaTextDemo.xlsx");
        }
    }
}
