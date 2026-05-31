using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Access the cells collection
            Cells cells = worksheet.Cells;

            // Define zero‑based row and column indices
            int rowIndex = 2;    // corresponds to row 3 in Excel (A3, B3, etc.)
            int columnIndex = 1; // corresponds to column B

            // Put a sample value into the target cell for demonstration
            cells[rowIndex, columnIndex].PutValue("Sample Text");

            // Access the cell using the zero‑based indices
            Cell targetCell = cells[rowIndex, columnIndex];

            // Read the cell value as a string (handles different data types)
            string cellValue = targetCell.StringValue;

            // Log the value to the console
            Console.WriteLine($"Value at cell ({rowIndex}, {columnIndex}) is: {cellValue}");

            // Optionally save the workbook to verify the cell content
            workbook.Save("ZeroBasedCellAccess.xlsx");
        }
    }
}