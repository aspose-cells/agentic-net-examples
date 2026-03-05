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

            // Access the cells collection of the worksheet
            Cells cells = worksheet.Cells;

            // Define row and column indexes (zero‑based)
            int rowIndex = 2;    // corresponds to the third row (e.g., "A3")
            int columnIndex = 1; // corresponds to the second column (e.g., "B")

            // Get the cell at the specified row and column using the Cells indexer
            Cell cell = cells[rowIndex, columnIndex];

            // Optionally put a value into the cell to demonstrate that it works
            cell.PutValue("Hello from row 2, column 1");

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