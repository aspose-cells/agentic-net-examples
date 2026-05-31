using System;
using System.Collections;
using System.Collections.Generic;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook and populate it with mixed data types
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        cells["A1"].PutValue("Hello");
        cells["B1"].PutValue(123);
        cells["A2"].PutValue("World");
        cells["B2"].PutValue(DateTime.Now);
        cells["C3"].PutValue("Aspose");

        // List to collect cells whose value type is string
        List<Cell> stringCells = new List<Cell>();

        // Enumerate all cells in the worksheet
        IEnumerator enumerator = cells.GetEnumerator();
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;

            // Filter only string‑type cells
            if (cell.Type == CellValueType.IsString)
            {
                stringCells.Add(cell);
            }
        }

        // Display the collected string cells
        Console.WriteLine($"String‑type cells found: {stringCells.Count}");
        foreach (Cell sc in stringCells)
        {
            Console.WriteLine($"{sc.Name}: {sc.StringValue}");
        }

        // Save the workbook (optional)
        workbook.Save("FilteredStrings.xlsx");
    }
}