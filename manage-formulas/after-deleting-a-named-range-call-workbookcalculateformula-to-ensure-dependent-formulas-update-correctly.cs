using Aspose.Cells;
using System;

class DeleteNamedRangeAndRecalculate
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate some data
        sheet.Cells["A1"].PutValue(10);
        sheet.Cells["A2"].PutValue(20);
        sheet.Cells["A3"].PutValue(30);

        // Add a named range "MyRange" that refers to A1:A3
        int nameIndex = workbook.Worksheets.Names.Add("MyRange");
        workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

        // Use the named range in a formula
        sheet.Cells["B1"].Formula = "=SUM(MyRange)";

        // Calculate formulas before removal
        workbook.CalculateFormula();

        // Delete the named range
        workbook.Worksheets.Names.Remove("MyRange");

        // Recalculate formulas after removal to update dependent cells
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("AfterNamedRangeRemoval.xlsx");
    }
}