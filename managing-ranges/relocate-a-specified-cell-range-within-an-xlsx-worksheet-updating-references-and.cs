using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate source range A1:B2 with data
        cells["A1"].PutValue("Header");
        cells["A2"].PutValue(10);
        cells["A3"].PutValue(20);
        cells["B2"].PutValue(100);
        cells["B3"].PutValue(200);

        // Add a formula that references the source range (A2:B3)
        // This will help us verify that references are updated after moving
        cells["C2"].Formula = "=SUM(A2:B3)";

        // Define the source area to move (A2:B3)
        // Using the CreateCellArea rule to build the CellArea object
        CellArea sourceArea = CellArea.CreateCellArea(1, 0, 2, 1); // rows 1-2, columns 0-1 (A2:B3)

        // Destination: move the range down 4 rows and right 3 columns (to D6:E7)
        int destRow = 5;    // zero‑based index for row 6
        int destColumn = 3; // zero‑based index for column D

        // Move the range using the Cells.MoveRange method (rule provided)
        cells.MoveRange(sourceArea, destRow, destColumn);

        // After moving, the formula in C2 should automatically adjust to reference the new location (D6:E7)
        // Verify by printing the updated formula
        Console.WriteLine("Updated formula in C2: " + cells["C2"].Formula);

        // Save the workbook (using the standard Save rule)
        workbook.Save("MovedRange.xlsx");
    }
}