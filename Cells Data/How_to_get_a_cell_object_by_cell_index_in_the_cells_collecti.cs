using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection of the worksheet
            Cells cells = worksheet.Cells;

            // Retrieve a cell by its zero‑based row and column indexes
            // Example: row index 2 (third row), column index 3 (fourth column) => cell D3
            Cell cell = cells[2, 3];

            // Put a value into the retrieved cell
            cell.PutValue("Hello from D3");

            // Save the workbook to an XLSX file
            workbook.Save("CellByIndexDemo.xlsx");
        }
    }
}