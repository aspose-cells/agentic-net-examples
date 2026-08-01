// Title: C# Aspose.Cells Example: Use IFERROR to Return a Default Value on Division‑by‑Zero
// Description: This C# demo creates a new Workbook, writes 10 to A1 and 0 to B1, applies the formula `=IFERROR(A1/B1, "Div/0!")` to C1, forces calculation, prints the result ("Div/0!"), and saves the file as IFERROR_DivisionByZero.xlsx.
// Keywords: Aspose.Cells IFERROR | C# Excel division by zero | handle #DIV/0! error | set Excel formula programmatically | calculate formulas Aspose.Cells | save workbook .NET
// Common Searches: Aspose.Cells IFERROR example C# | how to avoid #DIV/0! in Aspose.Cells | C# set default value for Excel division error | calculate workbook after adding formula Aspose.Cells | save Excel file with IFERROR using Aspose
// Developer Intent: Insert an IFERROR formula that supplies a fallback value when a division results in an error, evaluate the sheet, and persist the workbook.
// Use Cases: Display a custom message instead of #DIV/0! in generated reports. | Ensure calculations continue when input data may contain zero denominators. | Create reusable templates that automatically replace errors with safe defaults.
// AI Prompts: Show how to modify the IFERROR formula to return 0 instead of a text string. | Provide code that applies an IFERROR formula to every cell in column C based on column headers. | Explain strategies for handling multiple error‑prone formulas in a single worksheet with Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsIFERRORDemo
{
    // This C# demo creates a new Workbook, writes 10 to A1 and 0 to B1, applies the formula `=IFERROR(A1/B1, "Div/0!")` to C1, forces calculation, prints the result ("Div/0!"), and saves the file as IFERROR_DivisionByZero.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Put numerator and denominator values
            cells["A1"].PutValue(10);   // Numerator
            cells["B1"].PutValue(0);    // Denominator (will cause division by zero)

            // Set a formula that uses IFERROR to provide a default value when division by zero occurs
            // If A1/B1 results in an error, the formula returns the string "Div/0!"
            cells["C1"].Formula = "=IFERROR(A1/B1, \"Div/0!\")";

            // Calculate all formulas in the workbook (lifecycle rule: calculate)
            workbook.CalculateFormula();

            // Output the result of the IFERROR formula
            Console.WriteLine("Result in C1: " + cells["C1"].StringValue); // Expected: Div/0!

            // Save the workbook to a file (lifecycle rule: save)
            workbook.Save("IFERROR_DivisionByZero.xlsx");
        }
    }
}
