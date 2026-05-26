using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (lifecycle rule)
        Workbook wb = new Workbook();

        // First worksheet: will contain the dynamic array formula
        Worksheet dataSheet = wb.Worksheets[0];
        dataSheet.Name = "Data";

        // Populate source data that the dynamic array will depend on
        dataSheet.Cells["A1"].PutValue(1);
        dataSheet.Cells["A2"].PutValue(2);
        dataSheet.Cells["A3"].PutValue(3);

        // Set a dynamic array formula in B1 that spills vertically.
        // The formula uses SEQUENCE based on the count of values in A1:A3.
        // This will create a spill range B1:B3.
        Cell dynCell = dataSheet.Cells["B1"];
        dynCell.SetDynamicArrayFormula("=SEQUENCE(COUNTA(A1:A3))", new FormulaParseOptions(), true);

        // Refresh dynamic array formulas so the spill range is materialized (rule usage)
        wb.RefreshDynamicArrayFormulas(true);

        // Second worksheet: will sum the spilled range from the first sheet
        Worksheet summarySheet = wb.Worksheets.Add("Summary");

        // Use the spill operator (#) to reference the entire spilled range B1#
        // and calculate its sum.
        summarySheet.Cells["A1"].Formula = "=SUM(Data!B1#)";

        // Calculate all formulas in the workbook
        wb.CalculateFormula();

        // Save the workbook (lifecycle rule)
        wb.Save("DynamicArraySpillSum.xlsx");
    }
}