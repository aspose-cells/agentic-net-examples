using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

class InsertRowIntoListObjectAndExtendChart
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // ----- Populate initial data (Q1‑Q3) -----
        sheet.Cells["A1"].PutValue("Quarter");   // Header
        sheet.Cells["B1"].PutValue("Sales");     // Header
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["B2"].PutValue(100);
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["B3"].PutValue(150);
        sheet.Cells["A4"].PutValue("Q3");
        sheet.Cells["B4"].PutValue(200);

        // ----- Create a ListObject (table) that includes the data -----
        // Using the overload that takes start and end cell addresses
        int tableIndex = sheet.ListObjects.Add("A1", "B4", true);
        ListObject table = sheet.ListObjects[tableIndex];

        // ----- Add a chart that uses the table as its data source -----
        // The chart will automatically expand when the table grows
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Use structured references to the table columns
        // Table name defaults to "Table1" unless renamed
        chart.NSeries.Add("=Sheet1!Table1[Sales]", true);
        chart.NSeries.CategoryData = "=Sheet1!Table1[Quarter]";

        // ----- Insert a new row (Q4) into the ListObject -----
        // Calculate the row offset for the new row (first row after the current end)
        int newRowOffset = table.EndRow - table.StartRow + 1; // zero‑based offset within the table
        table.PutCellValue(newRowOffset, 0, "Q4");   // Quarter column
        table.PutCellValue(newRowOffset, 1, 250);   // Sales column

        // ----- Save the workbook -----
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}