// Title: Set a SORT dynamic array formula that spills and sum the spill on a different worksheet with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# create a SORT dynamic array that spills to a range | How to reference a spilled array using the # operator in another sheet with Aspose.Cells | Programmatically assign a dynamic array formula and calculate its sum across worksheets in C# | Calculate SUM of a SORT spill from Data sheet in Summary sheet using Aspose.Cells | Example of using SORT and spill reference in Aspose.Cells workbook
// Tags: dynamic array formula SORT Aspose.Cells C# | spill range reference using # operator | cross‑worksheet SUM of spilled array | programmatic formula assignment Aspose.Cells | calculate workbook formulas Aspose.Cells .NET

using System;
using Aspose.Cells;

// The example creates a workbook, renames the first sheet to "Data" and adds a "Summary" sheet. It fills A1:A10 on Data with numbers 10‑1, sets a dynamic array formula "=SORT(A1:A10)" in B1 that automatically spills to B1:B10, and on Summary places "=SUM(Data!B1#)" to sum the spilled values. After recalculating formulas, the workbook is saved as "DynamicArraySpill.xlsx".
class Program
{
    static void Main()
    {
        // Create a new workbook with a default worksheet
        Workbook workbook = new Workbook();

        // Rename the default sheet to "Data"
        Worksheet dataSheet = workbook.Worksheets[0];
        dataSheet.Name = "Data";

        // Add a second worksheet named "Summary"
        int summaryIndex = workbook.Worksheets.Add();
        Worksheet summarySheet = workbook.Worksheets[summaryIndex];
        summarySheet.Name = "Summary";

        // Populate column A in the Data sheet with sample numbers (A1:A10)
        for (int i = 0; i < 10; i++)
        {
            dataSheet.Cells[i, 0].PutValue(10 - i); // Example values 10,9,...,1
        }

        // Set a dynamic array formula in B1 that spills the sorted values of A1:A10
        // The spill range will be B1:B10 automatically
        dataSheet.Cells["B1"].Formula = "=SORT(A1:A10)";

        // In the Summary sheet, reference the spilled array using the # operator
        // and calculate the sum of the spilled values
        summarySheet.Cells["A1"].Formula = "=SUM(Data!B1#)";

        // Recalculate all formulas in the workbook
        workbook.CalculateFormula();

        // Save the workbook to a file
        workbook.Save("DynamicArraySpill.xlsx");
    }
}
