using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Load the existing workbook (or create a new one if the file does not exist)
        Workbook workbook;
        try
        {
            workbook = new Workbook("input.xlsx");
        }
        catch
        {
            workbook = new Workbook();
            workbook.Worksheets[0].Name = "Sheet1";
        }

        Worksheet worksheet = workbook.Worksheets[0];

        // Ensure there is at least one table (ListObject) in the worksheet
        ListObject table;
        if (worksheet.ListObjects.Count == 0)
        {
            // Add a table covering A1:B10 (including header row)
            int tableIndex = worksheet.ListObjects.Add(0, 0, 9, 1, true);
            table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "Table1";
        }
        else
        {
            table = worksheet.ListObjects[0];
        }

        // Insert values using row and column offsets relative to the table's start
        // Row offset 0 = header row, 1 = first data row, etc.
        table.PutCellValue(1, 0, 12345);               // First data row, first column
        table.PutCellValue(1, 1, "Sample Text");      // First data row, second column
        table.PutCellValue(2, 0, DateTime.Now);       // Second data row, first column

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}