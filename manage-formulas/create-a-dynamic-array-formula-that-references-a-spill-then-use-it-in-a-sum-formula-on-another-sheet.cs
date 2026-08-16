// Title: Aspose.Cells for .NET: Create a SEQUENCE dynamic array, spill it, and sum the spill range on another sheet
// Description: Shows how to build a workbook, add a "Data" sheet where B1 defines the sequence length, set a dynamic array formula =SEQUENCE(B1) that spills vertically, refresh the spill, then reference the spilled range with the # operator in a SUM formula on a "Summary" sheet, calculate all formulas, and save the file.
// Keywords: Aspose.Cells | .NET | C# | dynamic array formula | SEQUENCE function | spill range | # operator | SUM across worksheets | refresh dynamic array | Excel dynamic arrays | spreadsheet automation
// Common Searches: Aspose.Cells reference spilled array C# | SEQUENCE dynamic array Aspose.Cells example | SUM spilled range on another sheet Aspose.Cells | how to use # operator with Aspose.Cells | refresh dynamic array formulas Aspose.Cells
// Developer Intent: Demonstrate setting a SEQUENCE dynamic array, materializing its spill, and aggregating the spilled values from a different worksheet using the # operator in Aspose.Cells for .NET.
// Use Cases: Generate a vertical sequence whose size is driven by a cell value and let Excel spill the results automatically. | Refresh dynamic array formulas so the spill area becomes a concrete range that other formulas can address. | Reference a spilled array with the # operator in a SUM formula on another worksheet to calculate totals.
// AI Prompts: Provide C# code that creates a SEQUENCE dynamic array, refreshes the spill, and sums the spilled range on a separate sheet using Aspose.Cells. | Explain how the # operator works for referencing spilled array ranges across worksheets in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to build a workbook, add a "Data" sheet where B1 defines the sequence length, set a dynamic array formula =SEQUENCE(B1) that spills vertically, refresh the spill, then reference the spilled range with the # operator in a SUM formula on a "Summary" sheet, calculate all formulas, and save the file.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook wb = new Workbook();

        // -------------------------------------------------
        // Sheet 1: create a dynamic array that will spill
        // -------------------------------------------------
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";

        // Value in B1 determines the size of the sequence
        dataSheet.Cells["B1"].PutValue(5);

        // Set a dynamic array formula in A1 that spills vertically
        // The formula =SEQUENCE(B1) will produce 5 rows: 1,2,3,4,5
        dataSheet.Cells["A1"].SetDynamicArrayFormula("=SEQUENCE(B1)", new FormulaParseOptions(), true);

        // Refresh dynamic array formulas so the spill range is materialized
        wb.RefreshDynamicArrayFormulas(true);

        // -------------------------------------------------
        // Sheet 2: reference the spilled range using the # operator
        // -------------------------------------------------
        Worksheet summarySheet = wb.Worksheets.Add("Summary");

        // Sum the spilled range from Data!A1# (the # denotes the spill area)
        summarySheet.Cells["A1"].Formula = "=SUM(Data!A1#)";

        // Calculate all formulas in the workbook
        wb.CalculateFormula();

        // Save the workbook
        wb.Save("DynamicArraySpillSum.xlsx");
    }
}
