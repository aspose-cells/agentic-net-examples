// Title: Insert a row to shift a SEQUENCE dynamic array spill and recalculate formulas using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that writes a =SEQUENCE(5) formula in A1, saves the workbook, inserts a new row at position 2, recalculates, and saves the modified file. | Show how to preserve a spilled dynamic array when inserting rows in an Excel worksheet using Aspose.Cells, including updating the spill range after recalculation. | Provide a C# example that adds custom data to an inserted row while keeping the SEQUENCE array formula intact and correctly shifted.
// Common Searches: how to shift a spilled SEQUENCE array after inserting rows with Aspose.Cells .NET | C# Aspose.Cells insert row and keep dynamic array spill range updated | recalculate dynamic array formulas after row insertion using Aspose.Cells
// Tags: Aspose.Cells insert rows dynamic array spill | C# SEQUENCE function spill handling | recalculate formulas after row insertion Aspose.Cells | save workbook after shifting dynamic array | dynamic array spill range management .NET

using System;
using Aspose.Cells;

// // This program creates a new workbook, places a =SEQUENCE(5) formula in cell A1 (spilling to A1:A5), saves the file, inserts a row at Excel row 2 to move the spill to A2:A6, adds a custom value in the inserted row, recalculates formulas to update the spill range, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Place a dynamic array formula in A1 that spills into 5 rows (A1:A5)
        // The SEQUENCE function generates a vertical array of numbers 1 to 5
        sheet.Cells["A1"].Formula = "=SEQUENCE(5)";

        // Evaluate the formula so the spill range is populated
        workbook.CalculateFormula();

        // Save the workbook with the initial spill range
        workbook.Save("DynamicArraySpill.xlsx");

        // Insert a new row at index 1 (Excel row 2) to shift the spill range down
        // This will cause the spilled values to move down one row (A2:A6)
        sheet.Cells.InsertRows(1, 1);

        // Optionally add some data in the newly inserted row
        sheet.Cells["A2"].PutValue("Inserted Row");

        // Recalculate formulas after the row insertion to update the spill range
        workbook.CalculateFormula();

        // Save the workbook after shifting the spill range
        workbook.Save("DynamicArraySpill_Shifted.xlsx");
    }
}
