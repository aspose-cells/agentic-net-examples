using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class UpdateNamedRangeAfterTableResize
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate initial data for the table (2 columns, 5 rows)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Create a table (ListObject) covering the initial data range A1:B5
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 1, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Create a named range that refers to the table's data range
            int nameIndex = workbook.Worksheets.Names.Add("SalesData");
            Name namedRange = workbook.Worksheets.Names[nameIndex];
            // Set RefersTo without the leading '='
            namedRange.RefersTo = $"{sheet.Name}!{table.DataRange.RefersTo}";

            // Expand the table by adding 3 more rows of data
            for (int i = 5; i < 8; i++)
            {
                sheet.Cells[i, 0].PutValue($"Item {i + 1}");
                sheet.Cells[i, 1].PutValue((i + 1) * 10);
            }

            // Resize the table to include the new rows (endRow = 7)
            table.Resize(0, 0, 7, 1, true);

            // Update the named range to reference the new table range
            // DataRange.RefersTo returns something like "$A$1:$B$8"
            string newRefersTo = $"{table.DataRange.RefersTo}";
            // SetRefersTo expects the formula without leading '=', specify A1 style and global scope
            namedRange.SetRefersTo($"{sheet.Name}!{newRefersTo}", false, false);

            // Save the workbook
            workbook.Save("UpdatedNamedRange.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}