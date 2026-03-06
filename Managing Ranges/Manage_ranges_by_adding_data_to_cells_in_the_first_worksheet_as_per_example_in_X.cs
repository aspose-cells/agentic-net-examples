using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Populate some sample data in the worksheet
        cells["A1"].PutValue("A1");
        cells["A2"].PutValue("A2");
        cells["B1"].PutValue("B1");
        cells["B2"].PutValue("B2");

        // Create a range that covers A1:B2
        AsposeRange range = cells.CreateRange("A1", "B2");

        // Insert a row and a column to demonstrate that the range expands automatically
        cells.InsertRow(1);      // Inserts a row at index 1 (second row)
        cells.InsertColumn(1);   // Inserts a column at index 1 (second column)

        // Output the updated size of the range
        Console.WriteLine($"Range now has {range.RowCount} rows and {range.ColumnCount} columns");

        // Save the workbook to an XLSX file
        workbook.Save("ManagedRange.xlsx");
    }
}