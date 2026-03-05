using System;
using Aspose.Cells;

class AccessCellsDemo
{
    static void Main()
    {
        // Load an existing XLSX workbook
        Workbook workbook = new Workbook("input.xlsx"); // load rule

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Get the cells collection
        Cells cells = worksheet.Cells;

        // Access a cell by zero‑based row and column indexes (A1)
        Cell cellA1 = cells[0, 0];
        Console.WriteLine("A1 value: " + cellA1.StringValue);

        // Modify a cell by indexes (A2)
        cells[1, 0].PutValue("Updated Value");

        // Access a cell by its address name (B2)
        Cell cellB2 = cells["B2"];
        Console.WriteLine("B2 original value: " + cellB2.StringValue);

        // Change the value of B2
        cellB2.PutValue(12345);

        // Save the workbook back to XLSX format
        workbook.Save("output.xlsx"); // save rule
    }
}