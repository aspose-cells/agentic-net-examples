using System;
using Aspose.Cells;

namespace MergeAndFreezeHeaderDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Merge cells across the top row (e.g., A1 to E1)
            // Parameters: firstRow = 0, firstColumn = 0, totalRows = 1, totalColumns = 5
            worksheet.Cells.Merge(0, 0, 1, 5);

            // Set a value for the merged header cell
            worksheet.Cells[0, 0].PutValue("Merged Header");

            // Freeze the top row so the merged header stays visible while scrolling
            // Freeze at row index 1 (second row), column index 0, freezing 1 row and 0 columns
            worksheet.FreezePanes(1, 0, 1, 0);

            // Save the workbook to a file
            workbook.Save("MergedHeaderAndFrozenPane.xlsx");
        }
    }
}