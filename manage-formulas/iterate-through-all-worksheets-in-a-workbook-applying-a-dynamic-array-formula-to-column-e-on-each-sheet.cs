// Title: How to set a UNIQUE dynamic array formula in column E of every worksheet using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells apply UNIQUE(A:A) formula to column E on all worksheets | Iterate through workbook sheets and set a spill formula with Aspose.Cells | Programmatically add a dynamic array formula to each sheet in an Excel file using .NET | Set formula in cell E1 for every worksheet with Aspose.Cells library | How to assign a dynamic array formula across multiple worksheets in C#
// Tags: Aspose.Cells set cell formula C# | dynamic array UNIQUE formula Excel | iterate worksheets Aspose.Cells | apply spill formula column E | C# write formula to workbook | Aspose.Cells dynamic array across sheets

using Aspose.Cells;
using System;

// The code loads an existing workbook, defines the dynamic array formula =UNIQUE(A:A), loops through each worksheet, assigns the formula to cell E1 so the results spill down column E, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Define the dynamic array formula to be applied to column E
        // Example: returns unique values from column A
        string dynamicArrayFormula = "=UNIQUE(A:A)";

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Apply the dynamic array formula to the first cell of column E (E1)
            // The result will spill automatically into the cells below
            Cell targetCell = sheet.Cells["E1"];
            targetCell.Formula = dynamicArrayFormula;
        }

        // Save the updated workbook to a new file
        workbook.Save("output.xlsx");
    }
}
