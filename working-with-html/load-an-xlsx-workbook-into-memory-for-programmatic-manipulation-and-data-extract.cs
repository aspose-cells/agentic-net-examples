using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Path to the Excel file to be loaded
        string filePath = "input.xlsx";

        // Load the workbook into memory using the string constructor
        Workbook workbook = new Workbook(filePath);

        // Access the first worksheet in the workbook
        Worksheet worksheet = workbook.Worksheets[0];

        // Example: read the value from cell A1
        object a1Value = worksheet.Cells["A1"].Value;
        Console.WriteLine($"A1 value: {a1Value}");

        // Example: iterate over all used cells and display their contents
        int maxRow = worksheet.Cells.MaxDataRow;
        int maxCol = worksheet.Cells.MaxDataColumn;

        for (int row = 0; row <= maxRow; row++)
        {
            for (int col = 0; col <= maxCol; col++)
            {
                var cell = worksheet.Cells[row, col];
                if (cell.Value != null)
                {
                    Console.WriteLine($"Cell {cell.Name}: {cell.Value}");
                }
            }
        }

        // The workbook is now loaded in memory and can be manipulated further as needed.
    }
}