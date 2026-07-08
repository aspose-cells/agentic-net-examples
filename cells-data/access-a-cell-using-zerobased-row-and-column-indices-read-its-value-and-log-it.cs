using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // For demonstration, put a value into a cell at zero‑based row 0, column 0 (A1)
            worksheet.Cells[0, 0].PutValue("Hello Aspose!");

            // Access the same cell using zero‑based indices
            Cell cell = worksheet.Cells[0, 0];

            // Read the cell's value
            object cellValue = cell.Value;

            // Log the value along with its address
            Console.WriteLine($"Cell {cell.Name} (Row {cell.Row}, Column {cell.Column}) contains: {cellValue}");

            // Save the workbook (optional, just to illustrate the save lifecycle)
            workbook.Save("DemoOutput.xlsx");
        }
    }
}