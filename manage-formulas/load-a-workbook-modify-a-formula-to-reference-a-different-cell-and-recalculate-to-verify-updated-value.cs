// Title: C# – Load Excel workbook, modify a cell formula reference, recalculate with Aspose.Cells
// Description: Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, change the formula in cell A1 from referencing B1 to C1, trigger a full workbook recalculation, output the updated value, and save the result as a new file.
// Keywords: Aspose.Cells C# | modify Excel formula programmatically | recalculate workbook Aspose.Cells | change cell reference in formula | load and save Excel file .NET | Workbook.CalculateFormula example
// Common Searches: change formula cell reference Aspose.Cells C# | recalculate Excel workbook after formula edit | update Excel formula programmatically .NET | Aspose.Cells example modify cell formula | how to recalculate formulas with Aspose.Cells
// Developer Intent: Load an existing workbook, replace a formula’s cell reference, recalculate all formulas, and verify the new value programmatically.
// Use Cases: Automate updates to financial models when column positions shift. | Batch‑process reports to correct formula references after data restructuring. | Validate that formula changes produce expected results before publishing workbooks.
// AI Prompts: Write C# code with Aspose.Cells that changes the formula in B2 from "=D2*2" to "=E2*2", recalculates the workbook, and prints the new value. | Show how to iterate over a dictionary of old‑to‑new cell references, update each formula accordingly, and save the workbook. | Explain how to capture, log, and handle exceptions for the value returned after calling Workbook.CalculateFormula.

using System;
using Aspose.Cells;

// Demonstrates how to open an existing .xlsx file using Aspose.Cells for .NET, change the formula in cell A1 from referencing B1 to C1, trigger a full workbook recalculation, output the updated value, and save the result as a new file.
class Program
{
    static void Main()
    {
        // Load an existing workbook from disk
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet and its cells collection
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Assume cell A1 originally contains a formula that references B1.
        // Change the formula so it now references C1 instead.
        Cell cell = cells["A1"];
        cell.Formula = "=C1+10";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Output the updated value of the modified cell to verify the change
        Console.WriteLine("Updated value in A1: " + cell.Value);

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
