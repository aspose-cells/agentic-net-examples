// Title: Add a Q4 row to a ListObject table and have the linked column chart expand automatically using Aspose.Cells for .NET (C#)
// AI Prompts: Insert a new data row into an existing ListObject and let the associated column chart resize itself with Aspose.Cells in C#. | Use PutCellValue to add Q4 sales to a table and automatically refresh the chart that references the table. | Programmatically extend a structured table and have a column chart reflect the added row using the Aspose.Cells API.
// Common Searches: Aspose.Cells how to add a row to a ListObject and automatically update a linked chart in C# | C# insert data into table and extend column chart using structured references Aspose.Cells | extend chart range after inserting table row Aspose.Cells .NET example
// Tags: ListObject row insertion with chart auto‑extension | Aspose.Cells column chart dynamic data range | C# PutCellValue table update | structured reference chart source Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// // Demonstrates creating a workbook, defining a ListObject with Q1‑Q3 data, linking a column chart via structured references, inserting a Q4 row using PutCellValue, and saving the file where the chart automatically reflects the new data.
class InsertRowIntoListObject
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];

        // Populate initial data for Q1‑Q3
        ws.Cells["A1"].PutValue("Quarter");
        ws.Cells["B1"].PutValue("Sales");
        ws.Cells["A2"].PutValue("Q1");
        ws.Cells["B2"].PutValue(150);
        ws.Cells["A3"].PutValue("Q2");
        ws.Cells["B3"].PutValue(200);
        ws.Cells["A4"].PutValue("Q3");
        ws.Cells["B4"].PutValue(250);

        // Add a ListObject (table) that includes the data range
        int tableIdx = ws.ListObjects.Add("A1", "B4", true);
        ListObject table = ws.ListObjects[tableIdx];

        // Create a column chart that uses the table as its data source
        int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 5);
        Chart chart = ws.Charts[chartIdx];
        // Use structured references to the table columns
        chart.NSeries.Add($"={ws.Name}!{table.DisplayName}[Sales]", true);
        chart.NSeries.CategoryData = $"={ws.Name}!{table.DisplayName}[Quarter]";

        // Insert a new row for Q4 using PutCellValue.
        // Row offset is the current number of data rows (EndRow - StartRow)
        int newRowOffset = table.EndRow - table.StartRow;
        table.PutCellValue(newRowOffset, 0, "Q4");
        table.PutCellValue(newRowOffset, 1, 300);

        // Save the workbook; the chart automatically reflects the added row
        wb.Save("ListObjectChartExtended.xlsx", SaveFormat.Xlsx);
    }
}
