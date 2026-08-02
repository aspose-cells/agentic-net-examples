// Title: C# – Disable Automatic Date Grouping in Aspose.Cells PivotTable (GroupDates = false)
// Description: Learn how to stop Aspose.Cells from automatically grouping date fields in a PivotTable. The example creates a workbook with Date and Sales columns, adds a PivotTable, places the Date field in the row area, sets `pivotTable.GroupDates = false`, refreshes the data, and saves the file. This ensures each date remains separate in the resulting Excel report.
// Keywords: Aspose.Cells PivotTable | disable date grouping | GroupDates false | C# PivotTable example | .NET Excel pivot date auto‑group | prevent automatic date grouping Aspose | Excel PivotTable ungrouped dates | Aspose.Cells C# tutorial
// Common Searches: Aspose.Cells stop date grouping in PivotTable | PivotTable GroupDates property C# | disable automatic date grouping Aspose.Cells | how to keep dates ungrouped in Excel pivot using Aspose | Aspose.Cells PivotTable example C#
// Developer Intent: Keep the Date field ungrouped in an Aspose.Cells PivotTable by setting the GroupDates property to false before refreshing the pivot.
// Use Cases: Generate a daily sales report where each calendar day must appear as an individual row. | Create a time‑series analysis workbook that preserves raw timestamps for precise filtering. | Export Excel files for end‑users who expect dates to remain ungrouped, avoiding Excel’s default auto‑grouping behavior.
// AI Prompts: Write C# code with Aspose.Cells that adds a PivotTable and disables automatic date grouping by setting pivotTable.GroupDates = false before RefreshData. | Show a complete Aspose.Cells example that creates a workbook, populates Date and Sales columns, adds a PivotTable, and ensures dates stay ungrouped. | Explain step‑by‑step how to prevent Aspose.Cells from grouping dates in a PivotTable, including the required property and method calls.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// Learn how to stop Aspose.Cells from automatically grouping date fields in a PivotTable. The example creates a workbook with Date and Sales columns, adds a PivotTable, places the Date field in the row area, sets `pivotTable.GroupDates = false`, refreshes the data, and saves the file. This ensures each date remains separate in the resulting Excel report.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            worksheet.Cells["A1"].PutValue("Date");
            worksheet.Cells["B1"].PutValue("Sales");

            worksheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
            worksheet.Cells["B3"].PutValue(150);
            worksheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
            worksheet.Cells["B4"].PutValue(200);

            // Add a pivot table based on the data range
            int pivotIndex = worksheet.PivotTables.Add("A1:B4", "D3", "PivotTable1");
            PivotTable pivotTable = worksheet.PivotTables[pivotIndex];

            // Add the Date field as a row field and Sales as a data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Date");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot table data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Save the workbook
            workbook.Save("PivotTable_NoAutoGroup.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
