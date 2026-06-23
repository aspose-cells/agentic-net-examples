using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class CreateTableWithFilter
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a ListObject (table) covering the range A1:C10, with header row
        int tableIndex = worksheet.ListObjects.Add("A1", "C10", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Enable filter buttons on the header row of the table
        table.AutoFilter.ShowFilterButton = true;

        // Save the workbook
        workbook.Save("TableWithFilterButtons.xlsx");
    }
}