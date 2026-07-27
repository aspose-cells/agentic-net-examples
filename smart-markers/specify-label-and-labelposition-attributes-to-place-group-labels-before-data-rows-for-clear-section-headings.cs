using System;
using Aspose.Cells;
using Aspose.Cells.Tables;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data with two logical groups
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Amount");

        // First group header and its items
        worksheet.Cells["A2"].PutValue("Group 1");
        worksheet.Cells["B2"].PutValue(100);
        worksheet.Cells["A3"].PutValue("Item 1.1");
        worksheet.Cells["B3"].PutValue(40);
        worksheet.Cells["A4"].PutValue("Item 1.2");
        worksheet.Cells["B4"].PutValue(60);

        // Second group header and its items
        worksheet.Cells["A5"].PutValue("Group 2");
        worksheet.Cells["B5"].PutValue(200);
        worksheet.Cells["A6"].PutValue("Item 2.1");
        worksheet.Cells["B6"].PutValue(200);

        // Group rows for each logical section (0‑based indices)
        // Group rows 2‑4 (Group 1 header + its items)
        worksheet.Cells.GroupRows(1, 3);
        // Group rows 5‑6 (Group 2 header + its items)
        worksheet.Cells.GroupRows(4, 5);

        // Place the summary (totals) row *above* the detail rows
        // This makes the label appear before the data rows
        worksheet.Outline.SummaryRowBelow = false; // Outline.SummaryRowBelow property

        // Convert the range into a table so we can use a totals row as the group label
        int tableIndex = worksheet.ListObjects.Add(0, 0, 6, 1, true);
        ListObject table = worksheet.ListObjects[tableIndex];
        table.ShowTotals = true; // Enable totals row

        // Set the label that will appear in the totals row for the Amount column
        // This acts as the group label displayed before the grouped rows
        ListColumn amountColumn = table.ListColumns[1];
        amountColumn.TotalsRowLabel = "Group Total"; // ListColumn.TotalsRowLabel property

        // Save the workbook
        workbook.Save("GroupLabelsBeforeData.xlsx");
    }
}