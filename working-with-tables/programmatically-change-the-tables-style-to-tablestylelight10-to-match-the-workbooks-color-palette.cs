using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Add sample data for the table
        sheet.Cells["A1"].PutValue("Name");
        sheet.Cells["B1"].PutValue("Age");
        sheet.Cells["A2"].PutValue("John");
        sheet.Cells["B2"].PutValue(30);
        sheet.Cells["A3"].PutValue("Jane");
        sheet.Cells["B3"].PutValue(25);

        // Create a ListObject (table) that includes the data range
        int tableIdx = sheet.ListObjects.Add(0, 0, 2, 1, true);
        ListObject table = sheet.ListObjects[tableIdx];

        // Apply the built‑in style 'TableStyleLight10' to the table
        table.TableStyleType = TableStyleType.TableStyleLight10;

        // Save the workbook
        workbook.Save("TableStyleLight10.xlsx");
    }
}