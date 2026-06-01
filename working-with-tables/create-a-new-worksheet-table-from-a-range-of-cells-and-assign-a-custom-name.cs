using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;          // Needed for ListObject
using AsposeRange = Aspose.Cells.Range;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in the range A1:C4
            for (int row = 0; row < 4; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Create a Range object that covers the data (alias to avoid conflict with System.Range)
            AsposeRange dataRange = cells.CreateRange("A1", "C4");

            // Add a ListObject (table) based on the range
            int tableIndex = sheet.ListObjects.Add(
                dataRange.FirstRow,
                dataRange.FirstColumn,
                dataRange.FirstRow + dataRange.RowCount - 1,
                dataRange.FirstColumn + dataRange.ColumnCount - 1,
                true); // true indicates the range has a header row

            // Retrieve the created table
            ListObject table = sheet.ListObjects[tableIndex];

            // Assign a custom name to the table
            table.DisplayName = "CustomTable";

            // Save the workbook
            string outputPath = "TableWithCustomName.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}