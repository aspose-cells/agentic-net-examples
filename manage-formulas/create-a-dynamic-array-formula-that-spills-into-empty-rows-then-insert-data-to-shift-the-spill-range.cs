// Title: Insert a SEQUENCE dynamic array formula that spills into rows and shift the spill range by inserting a row using Aspose.Cells for .NET (C#)
// Common Searches: how to set a SEQUENCE formula that spills in Aspose.Cells C# | Aspose.Cells insert row to move spilled array values | C# example of dynamic array spill and row insertion with Aspose.Cells | shift spilled SEQUENCE results by adding a row in .NET workbook | Aspose.Cells create dynamic array and adjust spill range after insertion
// Tags: Aspose.Cells SEQUENCE dynamic array spill | insert rows shift spilled formula Aspose.Cells | C# workbook dynamic array manipulation | Excel dynamic array spill handling with Aspose | row insertion affecting spilled values .NET

using Aspose.Cells;
using System;

// The sample creates a new workbook, places a SEQUENCE(5,1,1,1) formula in cell A1 that spills five rows, saves the file, inserts a row at index 1 to move the spilled values down, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Set a dynamic array formula in A1 that spills 5 rows down
        // SEQUENCE generates a vertical array: 1,2,3,4,5
        sheet.Cells["A1"].Formula = "=SEQUENCE(5,1,1,1)";

        // Save the workbook with the spill formula
        workbook.Save("DynamicArraySpill.xlsx");

        // Insert a row at row 2 (index 1) to shift the spill range down
        // The spilled values will move down by one row
        sheet.Cells.InsertRows(1, 1); // Insert 1 row starting at row index 1

        // Save the workbook after shifting the spill range
        workbook.Save("DynamicArraySpillShifted.xlsx");
    }
}
