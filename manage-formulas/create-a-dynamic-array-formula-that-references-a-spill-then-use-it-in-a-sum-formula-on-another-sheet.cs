// Title: Create a SORT dynamic array spill and aggregate it on a different worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# reference a dynamic array spill (B1#) from another worksheet | How to sum a SORT spill range on a different sheet with Aspose.Cells | C# example to calculate formulas after creating a dynamic array spill in Aspose.Cells | Using SORT function and SUM across worksheets in Aspose.Cells | Create and save a workbook with a dynamic array spill and cross‑sheet SUM in C#
// Tags: SORT function spill handling Aspose.Cells | cross‑worksheet SUM of spilled array C# | dynamic array B1# reference Aspose.Cells | Workbook.CalculateFormula usage Aspose.Cells | export workbook to Excel with Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a workbook, fills column A with numbers 1‑10 on a sheet named "Data", adds a SORT dynamic array formula in B1 that spills its results (B1#), then adds a "Summary" sheet where A1 contains a SUM formula that references the spilled range Data!B1#. After calculating all formulas, the workbook is saved as DynamicArraySpillSum.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // First worksheet: populate data and create a dynamic array formula that spills
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Fill A1:A10 with sample numbers 1..10
        for (int i = 0; i < 10; i++)
        {
            dataSheet.Cells[i, 0].PutValue(i + 1);
        }

        // B1 contains a dynamic array formula (SORT) that spills its results into B1#
        dataSheet.Cells["B1"].Formula = "=SORT(A1:A10)";

        // Second worksheet: reference the spill range from the first sheet in a SUM formula
        Worksheet summarySheet = workbook.Worksheets.Add("Summary");
        // A1 sums all values that spilled from Data!B1#
        summarySheet.Cells["A1"].Formula = "=SUM(Data!B1#)";

        // Calculate formulas so that values are materialized
        workbook.CalculateFormula();

        // Save the workbook
        workbook.Save("DynamicArraySpillSum.xlsx");
    }
}
