// Title: C# – Retrieve the calculated value of cell E3 after formula evaluation with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, assign numeric values to A1, B1, and C1, set the formula "=A1+B1*C1" in E3, run workbook.CalculateFormula(), and read the resulting value from sheet.Cells["E3"].Value in a C# console application.
// Keywords: Aspose.Cells C# | Aspose.Cells .NET formula calculation | retrieve calculated cell value | workbook.CalculateFormula example | read cell E3 value | C# spreadsheet formula evaluation | Aspose.Cells get cell result
// Common Searches: Aspose.Cells get result of formula after CalculateFormula | C# read value of cell E3 with Aspose.Cells | How to evaluate formulas and fetch cell value in Aspose.Cells .NET | Example of calculating workbook formulas and retrieving a specific cell
// Developer Intent: Obtain the numeric result of cell E3 after the workbook formulas have been calculated.
// Use Cases: Display a computed total from a generated spreadsheet directly in a console or UI. | Pass a specific calculated figure (e.g., subtotal, tax) to another service or API. | Log or audit the outcome of financial or statistical formulas embedded in an automated report.
// AI Prompts: Generate C# code that sets a formula in a cell, runs workbook.CalculateFormula, and returns the computed value using Aspose.Cells. | Show how to retrieve multiple calculated cell values after calling CalculateFormula with Aspose.Cells for .NET. | Explain best practices for casting the Value property of a calculated cell to int, double, or string in Aspose.Cells.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, assign numeric values to A1, B1, and C1, set the formula "=A1+B1*C1" in E3, run workbook.CalculateFormula(), and read the resulting value from sheet.Cells["E3"].Value in a C# console application.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some cells with sample data
        sheet.Cells["A1"].PutValue(5);
        sheet.Cells["B1"].PutValue(10);
        sheet.Cells["C1"].PutValue(15);

        // Set a formula in E3 that uses the above cells
        // Example formula: =A1 + B1 * C1
        sheet.Cells["E3"].Formula = "=A1+B1*C1";

        // Calculate all formulas in the workbook
        workbook.CalculateFormula();

        // Retrieve the final calculated value of cell E3
        object finalValue = sheet.Cells["E3"].Value;

        // Display the result
        Console.WriteLine("Final calculated value of E3: " + finalValue);
    }
}
