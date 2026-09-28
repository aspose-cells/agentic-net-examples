// Title: Write a SUM formula to cell A1 and retrieve its text in cell D1 using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that places values in B1 and C1, writes a SUM formula to A1, uses =FORMULATEXT(A1) in D1 to capture the formula string, calculates all formulas, and saves the workbook. | Show how to extract an Excel formula as plain text by assigning the FORMULATEXT function to another cell and invoking workbook.CalculateFormula() in Aspose.Cells.
// Common Searches: Aspose.Cells C# how to use FORMULATEXT to get formula string from a cell | C# write SUM formula to a cell and read its text with Aspose.Cells | retrieve Excel formula as text in another cell using Aspose.Cells .NET | calculate formulas and extract formula text in Aspose.Cells workbook | save workbook after extracting formula text with Aspose.Cells C#
// Tags: Aspose.Cells write SUM formula | Aspose.Cells FORMULATEXT function | Aspose.Cells calculate workbook formulas | Aspose.Cells extract formula as text | C# save Excel workbook with Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a workbook, puts numbers in B1/C1, writes =SUM(B1,C1) to A1, stores the formula text in D1 via =FORMULATEXT(A1), calculates all formulas, and saves as FormulaTextDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Put some sample values that will be used in the formula
        cells["B1"].PutValue(5);
        cells["C1"].PutValue(10);

        // Write a formula to cell A1 (e.g., sum of B1 and C1)
        cells["A1"].Formula = "=SUM(B1,C1)";

        // Store the formula text of A1 into cell D1 using the Excel FORMULATEXT function
        cells["D1"].Formula = "=FORMULATEXT(A1)";

        // Calculate all formulas so that D1 contains the actual formula string
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("FormulaTextDemo.xlsx");
    }
}
