// Title: Insert a SEQUENCE dynamic array in one worksheet and sum its spilled range on another sheet using Aspose.Cells for .NET
// AI Prompts: Write C# with Aspose.Cells to place a SEQUENCE formula in Data!A2, then add a SUM formula on Summary!A1 that points to the spilled output of that array and evaluate the workbook. | Show how to reference a dynamic array spill from one sheet in a formula on a different sheet with Aspose.Cells and trigger calculation before saving.
// Common Searches: aspnet c# aspose.cells sum values from a SEQUENCE spill on another worksheet | how to use the # spill operator with SUM in an Aspose.Cells workbook | example of referencing a dynamic array created by SEQUENCE across sheets in Aspose.Cells | calculating formulas after inserting a dynamic array in Aspose.Cells .NET
// Tags: SEQUENCE function Aspose.Cells .NET | cross‑sheet spill reference Aspose.Cells | SUM over spilled array Aspose.Cells | formula calculation Aspose.Cells workbook | export workbook to XLSX Aspose.Cells

using Aspose.Cells;
using System;

// The program creates a workbook, adds a SEQUENCE dynamic array in Data!A2 that spills numbers 1‑5, creates a Summary sheet, uses the # spill operator in a SUM formula on Summary!A1 to total the spilled values, evaluates all formulas, and saves the file as DynamicArraySpillSum.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet: will contain the dynamic array formula
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Place a dynamic array formula in A2 that spills a sequence 1‑5 vertically
        // SEQUENCE(5,1,1,1) creates a 5‑row, 1‑column array starting at 1
        Cell arrayFormulaCell = dataSheet.Cells["A2"];
        arrayFormulaCell.Formula = "=SEQUENCE(5,1,1,1)";

        // Second worksheet: will sum the spilled range from the first sheet
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");

        // Use the spill operator (#) to reference the entire spilled range from Data!A2
        // SUM will add all numbers produced by the dynamic array
        Cell sumCell = summarySheet.Cells["A1"];
        sumCell.Formula = "=SUM(Data!A2#)";

        // Evaluate all formulas so the workbook contains the calculated values
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DynamicArraySpillSum.xlsx");
    }
}
