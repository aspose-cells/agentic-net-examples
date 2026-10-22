// Title: Create a spilling SEQUENCE dynamic array formula and shift its range by inserting rows with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells create a SEQUENCE formula that spills into a vertical range | how to move a spilled dynamic array down by inserting rows using Aspose.Cells | recalculate formulas after inserting rows in an Aspose.Cells workbook | save workbook before and after shifting a spilled array formula with Aspose.Cells
// Tags: Aspose.Cells set SEQUENCE dynamic array formula | insert rows to shift spilled array range Aspose.Cells | recalculate formulas after row insertion C# | save workbook with updated spill Aspose.Cells | dynamic array spill handling Aspose.Cells .NET

using Aspose.Cells;
using System;

// The example demonstrates how to use Aspose.Cells for .NET to set a SEQUENCE formula in cell A1 that spills into a vertical range, calculate and save the initial workbook, insert a row to shift the spilled range down, add data to the inserted row, recalculate the formulas, and save the workbook with the updated spill range.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a dynamic array formula in A1 that spills 5 rows (A1:A5)
        // SEQUENCE generates a vertical array of numbers 1 to 5
        sheet.Cells["A1"].Formula = "=SEQUENCE(5,1,1,1)";

        // Calculate the formula so the spill range is populated
        workbook.CalculateFormula();

        // Save the workbook with the initial spill
        workbook.Save("DynamicArraySpill.xlsx");

        // Insert a new row at row 2 (index 1) to shift the spill range down
        // This will move the original spill from A1:A5 to A2:A6
        sheet.Cells.InsertRows(1, 1); // Insert one row above the original A2

        // Add some data in the inserted row to illustrate the shift
        sheet.Cells["A2"].PutValue("Inserted Row");

        // Recalculate formulas after the row insertion
        workbook.CalculateFormula();

        // Save the workbook after shifting the spill range
        workbook.Save("DynamicArraySpill_Shifted.xlsx");
    }
}
