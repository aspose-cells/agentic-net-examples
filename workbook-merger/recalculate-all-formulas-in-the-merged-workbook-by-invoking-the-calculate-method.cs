using System;
using Aspose.Cells;

class MergeAndCalculate
{
    static void Main()
    {
        // Load the primary workbook
        Workbook workbook = new Workbook("Book1.xlsx");

        // Load the workbook to be merged
        Workbook workbookToMerge = new Workbook("Book2.xlsx");

        // Copy each worksheet from the second workbook into the first workbook
        foreach (Worksheet sheet in workbookToMerge.Worksheets)
        {
            // Add a copy of the worksheet to the target workbook
            workbook.Worksheets.AddCopy(sheet.Name);
        }

        // After merging, recalculate all formulas in the combined workbook
        workbook.CalculateFormula();

        // Save the merged workbook with updated formula results
        workbook.Save("MergedResult.xlsx");
    }
}