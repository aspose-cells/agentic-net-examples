using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    class RetrieveCellByIndex
    {
        static void Main()
        {
            // Load an existing XLSX file (replace with your actual file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet in the workbook
            Worksheet worksheet = workbook.Worksheets[0];

            // Define zero‑based row and column indexes
            int rowIndex = 2;    // Example: third row (A3)
            int columnIndex = 1; // Example: second column (B)

            // Retrieve the cell using the Cells indexer (row, column)
            Cell cell = worksheet.Cells[rowIndex, columnIndex];

            // Output the cell's string representation (will be empty if the cell is null)
            Console.WriteLine($"Cell at [{rowIndex}, {columnIndex}] contains: \"{cell.StringValue}\"");

            // Optionally modify the cell value
            cell.PutValue("Updated Value");

            // Save the workbook to a new file (replace with your desired output path)
            workbook.Save("output.xlsx");
        }
    }
}