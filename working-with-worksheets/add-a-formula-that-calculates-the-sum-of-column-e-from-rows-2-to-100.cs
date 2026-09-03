// Title: Add a SUM formula for cells E2 through E100 and place the result in F1 using Aspose.Cells for .NET (C#)
// AI Prompts: Use Aspose.Cells in C# to assign the formula "=SUM(E2:E100)" to cell F1 and then evaluate the workbook. | Create a new workbook, insert a SUM formula for the range E2:E100 into cell F1, calculate the formula, and save the file as XLSX. | Write C# code that programmatically totals column E rows 2‑100, stores the total in cell F1, and persists the workbook with Aspose.Cells.
// Common Searches: how to add a SUM formula to a specific range in Excel using Aspose.Cells C# | Aspose.Cells set formula '=SUM(E2:E100)' in cell F1 and calculate it | C# program to calculate total of column E rows 2 to 100 with Aspose.Cells and save workbook
// Tags: Aspose.Cells set cell formula C# | SUM formula insertion Excel .NET | calculate column total Aspose.Cells | evaluate workbook formulas C# | save workbook as XLSX Aspose.Cells

using System;
using Aspose.Cells;

// // Creates a workbook, inserts a SUM formula for E2:E100 into F1, evaluates the formula, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Get the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Place the SUM formula in cell F1 (you can choose any cell)
        Cell formulaCell = sheet.Cells["F1"];
        formulaCell.Formula = "=SUM(E2:E100)";

        // Optional: calculate the formula so the result is stored
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("output.xlsx");
    }
}
