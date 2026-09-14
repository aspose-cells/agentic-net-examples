// Title: Group numeric pivot field values into custom intervals (0‑20, 20‑40, 40‑60, 60‑80, 80‑100) using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that builds a pivot table, adds a numeric column as a row field, and groups it into 0‑20, 20‑40, 40‑60, 60‑80, and 80‑100 intervals using the PivotField.GroupBy method. | Show how to change the start, end, and interval arguments of PivotField.GroupBy to create custom range groups for a numeric field in an Excel pivot table with Aspose.Cells. | Demonstrate refreshing and calculating a pivot table after applying custom numeric grouping, then saving the workbook to a file.
// Common Searches: Aspose.Cells C# group pivot table numeric field into custom ranges | PivotField.GroupBy start end interval parameters example Aspose.Cells | Create Excel pivot table with interval grouping for amount column using Aspose.Cells .NET | Add the same numeric column as data and row field in Aspose.Cells pivot table | Refresh and calculate pivot table after numeric grouping Aspose.Cells
// Tags: Aspose.Cells pivot field custom range grouping | C# PivotField.GroupBy numeric intervals | Excel pivot table interval grouping Aspose.Cells | Aspose.Cells refresh calculate pivot after grouping | C# generate pivot table with grouped numeric values

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// The example creates a workbook, fills it with category and amount data, adds a pivot table, places the Category field as rows and the Amount field as both data and row fields, then uses PivotField.GroupBy to segment the numeric Amount row field into five custom intervals (0‑20, 20‑40, 40‑60, 60‑80, 80‑100). After refreshing and calculating the pivot, the workbook is saved as GroupedNumericPivot.xlsx.
class GroupNumericPivotField
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data: a category column and a numeric amount column
        sheet.Cells["A1"].Value = "Category";
        sheet.Cells["B1"].Value = "Amount";

        string[] categories = { "A", "A", "B", "B", "C", "C", "C", "D", "D", "E" };
        double[] amounts = { 5, 12, 18, 25, 33, 45, 58, 62, 77, 90 };

        for (int i = 0; i < categories.Length; i++)
        {
            sheet.Cells[i + 1, 0].Value = categories[i];
            sheet.Cells[i + 1, 1].Value = amounts[i];
        }

        // Create a pivot table based on the data range A1:B11 and place it at D3
        int pivotIndex = sheet.PivotTables.Add("A1:B11", "D3", "SalesPivot");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];

        // Add Category as a row field and Amount as a data field
        pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
        pivotTable.AddFieldToArea(PivotFieldType.Data, "Amount");

        // Also add Amount as a row field so we can group its numeric values
        int amountRowFieldIdx = pivotTable.AddFieldToArea(PivotFieldType.Row, "Amount");
        PivotField amountRowField = pivotTable.RowFields[amountRowFieldIdx];

        // Group numeric values into custom ranges:
        // 0‑20, 20‑40, 40‑60, 60‑80, 80‑100
        // start = 0, end = 100, interval = 20, newField = true (creates a new grouped field)
        amountRowField.GroupBy(0, 100, 20, true);

        // Refresh data and calculate the pivot table to apply the grouping
        pivotTable.RefreshData();
        pivotTable.CalculateData();

        // Save the workbook with the grouped pivot table
        workbook.Save("GroupedNumericPivot.xlsx");
    }
}
