using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    public class GetCellByRowColumnDemo
    {
        public static void Run()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Get the Cells collection of the worksheet
            Cells cells = worksheet.Cells;

            // Define row and column indexes (zero‑based)
            int rowIndex = 2;    // corresponds to Excel row 3
            int columnIndex = 4; // corresponds to Excel column E

            // Retrieve the cell at the specified row and column using the Cells indexer
            Cell cell = cells[rowIndex, columnIndex];

            // Put a value into the retrieved cell
            cell.PutValue("Hello from row 3, column E");

            // Save the workbook to an XLSX file
            workbook.Save("GetCellByRowColumnDemo.xlsx");
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            GetCellByRowColumnDemo.Run();
        }
    }
}