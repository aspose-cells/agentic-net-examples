using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: a text column and a numeric column
        worksheet.Cells["A1"].PutValue("Item");
        worksheet.Cells["B1"].PutValue("Quantity");
        worksheet.Cells["A2"].PutValue("Apple");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("Banana");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("Apple");
        worksheet.Cells["B4"].PutValue(15);

        // Add a ListObject (table) that includes the data range
        int tableIndex = worksheet.ListObjects.Add("A1", "B4", true);
        ListObject table = worksheet.ListObjects[tableIndex];

        // Enable the totals row for the table
        table.ShowTotals = true;

        // Set the totals calculation for the text column (first column) to Count
        table.ListColumns[0].TotalsCalculation = TotalsCalculation.Count;

        // Optionally set a label for the totals row of that column
        table.ListColumns[0].TotalsRowLabel = "Count";

        // Save the workbook to a file
        workbook.Save("ListObjectTotalsCountDemo.xlsx");
    }
}