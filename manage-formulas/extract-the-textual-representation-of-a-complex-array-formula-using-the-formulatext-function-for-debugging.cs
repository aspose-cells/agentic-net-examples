// Title: Extract array formula text with FORMULATEXT in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills two ranges, sets a 3‑row array formula "=A1:A3*B1:B3" in C1:C3, uses the Excel FORMULATEXT function in D1 to retrieve the exact formula string, calculates, prints the result, and saves the file.
// Keywords: Aspose.Cells | C# | .NET | array formula | FORMULATEXT | formula text extraction | debug Excel formulas | SetArrayFormula | FormulaText function
// Common Searches: Aspose.Cells get formula text of array formula | C# FORMULATEXT example with Aspose.Cells | retrieve array formula string in .NET | debug complex array formulas Aspose.Cells | how to use SetArrayFormula and FORMULATEXT
// Developer Intent: Obtain the exact textual representation of a multi‑cell array formula for debugging or documentation using Aspose.Cells.
// Use Cases: Log the precise formula applied to an array range to verify implementation. | Compare retrieved formula text with expected patterns during automated tests. | Generate audit reports that list all array formulas present in a workbook.
// AI Prompts: Show C# code that extracts the text of a multi‑cell array formula using Aspose.Cells and FORMULATEXT. | Explain how to debug complex array formulas in Aspose.Cells by retrieving their string representation. | Provide an example of handling cases where FORMULATEXT returns an empty string for a non‑array cell.

using System;
using Aspose.Cells;

// Creates a workbook, fills two ranges, sets a 3‑row array formula "=A1:A3*B1:B3" in C1:C3, uses the Excel FORMULATEXT function in D1 to retrieve the exact formula string, calculates, prints the result, and saves the file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate some sample data that will be used by the array formula
        worksheet.Cells["A1"].PutValue(1);
        worksheet.Cells["A2"].PutValue(2);
        worksheet.Cells["A3"].PutValue(3);
        worksheet.Cells["B1"].PutValue(4);
        worksheet.Cells["B2"].PutValue(5);
        worksheet.Cells["B3"].PutValue(6);

        // Set a complex array formula (e.g., element‑wise multiplication of two ranges)
        // The formula will occupy a 3‑row by 1‑column range starting at C1
        worksheet.Cells["C1"].SetArrayFormula("=A1:A3*B1:B3", 3, 1);

        // Use the Excel FORMULATEXT function to retrieve the textual representation
        // of the array formula from the first cell of the array (C1)
        worksheet.Cells["D1"].Formula = "=FORMULATEXT(C1)";

        // Calculate all formulas so that FORMULATEXT returns the actual text
        workbook.CalculateFormula();

        // Output the extracted formula text for debugging purposes
        Console.WriteLine("Extracted array formula: " + worksheet.Cells["D1"].StringValue);

        // Save the workbook (optional, demonstrates lifecycle compliance)
        workbook.Save("ArrayFormulaDebug.xlsx");
    }
}
