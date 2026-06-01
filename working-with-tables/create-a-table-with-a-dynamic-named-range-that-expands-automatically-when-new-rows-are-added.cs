using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

class DynamicTableWithNamedRange
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data with headers
            cells["A1"].PutValue("ID");
            cells["B1"].PutValue("Name");
            cells["A2"].PutValue(1);
            cells["B2"].PutValue("Alice");
            cells["A3"].PutValue(2);
            cells["B3"].PutValue("Bob");

            // Create a ListObject (table) that covers the data range A1:B3
            int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "MyTable"; // friendly name for the table

            // Define a dynamic named range that refers to the whole table.
            // The structured reference "MyTable[#All]" expands automatically as rows are added.
            int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"={sheet.Name}!{table.DisplayName}[#All]";

            // Determine the next row index after the current table data
            int lastDataRow = table.DataRange.FirstRow + table.DataRange.RowCount - 1;
            int newRow = lastDataRow + 1;

            // Write new data (the table expands automatically when saved)
            cells[newRow, 0].PutValue(3);          // ID column
            cells[newRow, 1].PutValue("Charlie"); // Name column

            // Save the workbook
            string outputPath = "DynamicTableNamedRange.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}