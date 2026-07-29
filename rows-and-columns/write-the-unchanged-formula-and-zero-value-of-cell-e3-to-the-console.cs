// Title: Aspose.Cells C# – Print raw formula and calculated value of cell E3 (zero result)
// Description: Creates a new Workbook, assigns the formula "=SUM(A1:A2)" to cell E3, writes the unchanged formula string to the console, evaluates all formulas, and then outputs the cell's DoubleValue, which is 0 when A1 and A2 are empty.
// Keywords: Aspose.Cells | C# | raw formula | cell formula text | DoubleValue | calculate formulas | console output | E3 | SUM function | zero result | Worksheet | Workbook
// Common Searches: Aspose.Cells get raw formula string | How to read unchanged formula from a cell in Aspose.Cells | Retrieve calculated numeric value of a formula with Aspose.Cells | Print cell formula and value to console using Aspose.Cells C# | Why does SUM return zero in Aspose.Cells when source cells are empty
// Developer Intent: Show the original formula text and its evaluated numeric result for cell E3.
// Use Cases: Debug spreadsheet calculations by logging both the formula and its result. | Create a simple audit report that lists formulas alongside computed values. | Confirm that a SUM formula yields zero when referenced cells contain no data.
// AI Prompts: Generate C# code with Aspose.Cells that sets a formula in E3, prints the raw formula, calculates the workbook, and displays the DoubleValue. | Explain how to obtain the Formula property and the DoubleValue property of a cell in Aspose.Cells and why they differ. | Provide a step‑by‑step guide to log a cell's original formula and its evaluated value to the console using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Creates a new Workbook, assigns the formula "=SUM(A1:A2)" to cell E3, writes the unchanged formula string to the console, evaluates all formulas, and then outputs the cell's DoubleValue, which is 0 when A1 and A2 are empty.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Set a formula in cell E3 (the formula is stored unchanged)
            sheet.Cells["E3"].Formula = "=SUM(A1:A2)";

            // Output the unchanged formula text
            Console.WriteLine("Unchanged formula in E3: " + sheet.Cells["E3"].Formula);

            // Calculate the workbook so that formula results are evaluated
            workbook.CalculateFormula();

            // Output the numeric value after calculation
            Console.WriteLine("Current numeric value in E3: " + sheet.Cells["E3"].DoubleValue);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
