using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GetCellByIndexDemo
    {
        public static void Run()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the cells collection of the worksheet
            Cells cells = worksheet.Cells;

            // Retrieve a cell by its zero‑based row and column indexes
            // Example: row index 2 (third row), column index 3 (fourth column) => cell D3
            Cell cell = cells[2, 3];
            cell.PutValue("Hello from D3");

            // Optionally, demonstrate accessing another cell using the same indexer
            Cell anotherCell = cells[0, 0]; // Cell A1
            anotherCell.PutValue(12345);

            // Save the workbook to an XLSX file
            workbook.Save("GetCellByIndexDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GetCellByIndexDemo.Run();
        }
    }
}