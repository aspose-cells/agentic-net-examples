using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeInfoDemo
{
    class Program
    {
        static void Main()
        {
            // Load an existing XLSX workbook (replace with your file path)
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Create a range using its address (example: B2:D5)
            AsposeRange range = cells.CreateRange("B2", "D5");

            // Retrieve basic range information
            string address = range.Address;                     // e.g., B2:D5
            int rowCount = range.RowCount;                     // number of rows in the range
            int columnCount = range.ColumnCount;               // number of columns in the range
            int cellCount = rowCount * columnCount;            // total cells in the range

            // Get an offset range (shift 1 row down and 2 columns to the right)
            AsposeRange offsetRange = range.GetOffset(1, 2);
            string offsetAddress = offsetRange.Address;        // address of the offset range

            // Get the entire column(s) that intersect the range
            AsposeRange entireColumnRange = range.EntireColumn;
            string entireColumnAddress = entireColumnRange.Address; // e.g., B:D

            // Get the entire row(s) that intersect the range
            AsposeRange entireRowRange = range.EntireRow;
            string entireRowAddress = entireRowRange.Address; // e.g., 2:5

            // Output the collected information
            Console.WriteLine($"Original Range Address: {address}");
            Console.WriteLine($"Rows: {rowCount}, Columns: {columnCount}, Cells: {cellCount}");
            Console.WriteLine($"Offset Range Address (1 row, 2 cols): {offsetAddress}");
            Console.WriteLine($"Entire Column(s) Address: {entireColumnAddress}");
            Console.WriteLine($"Entire Row(s) Address: {entireRowAddress}");

            // Optionally, save the workbook after any modifications
            workbook.Save("output.xlsx");
        }
    }
}