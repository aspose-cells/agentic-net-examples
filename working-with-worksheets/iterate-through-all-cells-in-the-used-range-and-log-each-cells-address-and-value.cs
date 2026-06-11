using System;
using System.Collections;
using Aspose.Cells;

public class IterateAllCells
{
    public static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Optional: add some sample data to demonstrate iteration
        cells["A1"].PutValue("Hello");
        cells["B2"].PutValue(123);
        cells["C3"].PutValue(DateTime.Now);

        // Obtain an enumerator that iterates over all cells containing data
        IEnumerator enumerator = cells.GetEnumerator();

        // Iterate through each cell, logging its address (Name) and value
        while (enumerator.MoveNext())
        {
            Cell cell = (Cell)enumerator.Current;
            Console.WriteLine($"{cell.Name}: {cell.Value}");
        }

        // Save the workbook to a file
        workbook.Save("IterateAllCells.xlsx");
    }
}