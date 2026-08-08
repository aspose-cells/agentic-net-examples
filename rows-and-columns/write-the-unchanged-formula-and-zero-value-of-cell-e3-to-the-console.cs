// Title: Print cell E3 formula and its default zero value to console with Aspose.Cells for .NET (C#)
// Description: Creates a new workbook, assigns the formula "=A1+B1" to cell E3, and writes both the stored formula string and the cell's current value (still 0 because CalculateFormula() hasn't been called) to the console.
// Keywords: Aspose.Cells | C# | .NET | cell formula retrieval | default cell value | uncalculated workbook | console output | E3 cell | read formula without calculation | Excel automation
// Common Searches: Aspose.Cells get cell formula without calculating | C# print cell value before calculation Aspose.Cells | How to display stored formula and default value in Aspose.Cells | Read formula from Excel cell using Aspose.Cells .NET | Console output of cell formula Aspose.Cells C#
// Developer Intent: Output the formula stored in E3 and its unchanged numeric value (0) without triggering any calculation.
// Use Cases: Debugging spreadsheet logic by logging formulas and their placeholder values before evaluation. | Creating an audit trail that records each cell's expression and initial value for version control. | Generating documentation that lists formulas alongside their default values for review.
// AI Prompts: Show how to retrieve and display a cell's formula and its current value without calculating it using Aspose.Cells for .NET. | Provide a C# example that writes the stored formula and default zero value of a specific cell to a log or console. | Explain how to output the unchanged value of a formula‑assigned cell while keeping the workbook in an uncalculated state.

using System;
using Aspose.Cells;

// Creates a new workbook, assigns the formula "=A1+B1" to cell E3, and writes both the stored formula string and the cell's current value (still 0 because CalculateFormula() hasn't been called) to the console.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Set a formula in cell E3 (the formula itself is stored, but not yet calculated)
        worksheet.Cells["E3"].Formula = "=A1+B1";

        // The cell's value is still the default (zero) because we haven't called CalculateFormula()
        Console.WriteLine("Formula in E3: " + worksheet.Cells["E3"].Formula);
        Console.WriteLine("Value in E3 (unchanged): " + worksheet.Cells["E3"].Value);
    }
}
