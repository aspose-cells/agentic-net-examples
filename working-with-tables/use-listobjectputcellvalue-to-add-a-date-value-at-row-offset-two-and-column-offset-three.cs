using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AsposeCellsDemo
{
    public class ListObjectPutDateExample
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data that will become the table (including headers)
                worksheet.Cells["A1"].PutValue("ID");
                worksheet.Cells["B1"].PutValue("Name");
                worksheet.Cells["C1"].PutValue("Amount");
                worksheet.Cells["D1"].PutValue("Date"); // This column will receive the date via PutCellValue

                // Sample data rows
                worksheet.Cells["A2"].PutValue(1);
                worksheet.Cells["B2"].PutValue("Alice");
                worksheet.Cells["C2"].PutValue(100);

                worksheet.Cells["A3"].PutValue(2);
                worksheet.Cells["B3"].PutValue("Bob");
                worksheet.Cells["C3"].PutValue(200);

                // Add a ListObject (table) covering the range A1:D3, with headers
                int startRow = 0;      // Row index for "A1"
                int startColumn = 0;   // Column index for "A1"
                int endRow = 2;        // Row index for "D3"
                int endColumn = 3;     // Column index for "D3"
                int tableIndex = worksheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
                ListObject table = worksheet.ListObjects[tableIndex];

                // Insert a date value at row offset 2 (third data row) and column offset 3 (fourth column)
                // The table will automatically expand by one row
                DateTime dateToInsert = new DateTime(2023, 12, 31);
                table.PutCellValue(2, 3, dateToInsert);

                // Format the newly added date cell
                Style dateStyle = workbook.CreateStyle();
                dateStyle.Number = 14; // Built‑in date format
                int newRowIndex = startRow + 1 + 2; // header row + offset
                worksheet.Cells[newRowIndex, 3].SetStyle(dateStyle);

                // Save the workbook
                workbook.Save("ListObjectPutDateDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ListObjectPutDateExample.Run();
        }
    }
}