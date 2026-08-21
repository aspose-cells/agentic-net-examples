// Title: C# Aspose.Cells Example: Use IFERROR to Return a Default Value on Division‑by‑Zero
// Description: This sample creates a workbook, writes 100 to A1 and 0 to B1, then sets C1 to the formula =IFERROR(A1/B1, "N/A"). The IFERROR function replaces the #DIV/0! error with "N/A", the workbook calculates the formula, prints the result, and saves the file as IFERRORDemo.xlsx.
// Keywords: Aspose.Cells IFERROR C# | handle division by zero Aspose.Cells | Excel IFERROR formula .NET | conditional formula Aspose.Cells | C# set Excel formula with Aspose | error handling in Excel using Aspose.Cells | Aspose.Cells calculate formulas | save workbook Aspose.Cells
// Common Searches: Aspose.Cells IFERROR example | C# avoid #DIV/0! error with Aspose.Cells | how to use IFERROR in Aspose.Cells .NET | set default value for Excel errors using Aspose | calculate formulas after applying IFERROR in C#
// Developer Intent: Insert an IFERROR formula so that a division‑by‑zero operation returns a specified fallback value.
// Use Cases: Replace #DIV/0! with "N/A" when computing ratios between cells. | Provide a safe default for any calculation that might generate an error, keeping reports clean. | Automate error‑proof Excel report generation by applying IFERROR to critical formulas.
// AI Prompts: Generate C# code with Aspose.Cells that applies =IFERROR(A1/B1, "N/A") and reads the result as a string. | Show how to apply IFERROR to a range of cells, calculate all formulas, and save the workbook using Aspose.Cells for .NET. | Explain how to customize the fallback value in an IFERROR formula when using Aspose.Cells in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsIFERRORDemo
{
    // This sample creates a workbook, writes 100 to A1 and 0 to B1, then sets C1 to the formula =IFERROR(A1/B1, "N/A"). The IFERROR function replaces the #DIV/0! error with "N/A", the workbook calculates the formula, prints the result, and saves the file as IFERRORDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set numerator and denominator values
            cells["A1"].PutValue(100);   // Numerator
            cells["B1"].PutValue(0);     // Denominator (will cause division by zero)

            // Apply IFERROR to return a default value when division by zero occurs
            // If A1/B1 results in an error, the cell will display "N/A"
            cells["C1"].Formula = "=IFERROR(A1/B1, \"N/A\")";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the result of the conditional formula
            Console.WriteLine("Result in C1: " + cells["C1"].StringValue);

            // Save the workbook (optional)
            workbook.Save("IFERRORDemo.xlsx");
        }
    }
}
