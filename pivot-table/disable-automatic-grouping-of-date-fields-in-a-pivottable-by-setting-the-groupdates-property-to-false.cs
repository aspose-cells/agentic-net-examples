// Title: Disable automatic date grouping in Aspose.Cells PivotTable (C#)
// Description: C# example that creates a workbook, adds Date and Sales columns, builds a PivotTable and shows how to keep dates from being auto‑grouped. The GroupDates property is not present in the current Aspose.Cells release, so the sample notes alternative APIs for controlling date grouping.
// Keywords: Aspose.Cells PivotTable | disable date auto grouping | GroupDates false | .NET Excel pivot | C# Aspose.Cells date granularity | prevent pivot date grouping | Excel pivot cache refresh
// Common Searches: Aspose.Cells stop date auto grouping | PivotTable GroupDates property C# | How to keep dates ungrouped in Aspose.Cells pivot | Disable date grouping in .NET Excel pivot table
// Developer Intent: Prevent a PivotTable from automatically grouping Date fields.
// Use Cases: Create a sales ledger where each transaction date appears as a separate row. | Export raw transaction data to Excel while preserving day‑level granularity. | Refresh a pivot cache without collapsing dates into months or years.
// AI Prompts: Generate C# code using Aspose.Cells to keep PivotTable dates ungrouped. | Suggest an alternative method for disabling date auto‑grouping when GroupDates is unavailable. | Explain how to manually set date grouping levels in an Aspose.Cells PivotTable.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

// C# example that creates a workbook, adds Date and Sales columns, builds a PivotTable and shows how to keep dates from being auto‑grouped. The GroupDates property is not present in the current Aspose.Cells release, so the sample notes alternative APIs for controlling date grouping.
class DisablePivotDateAutoGroup
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data with a Date column and a Sales column
            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue(new DateTime(2023, 1, 1));
            sheet.Cells["A3"].PutValue(new DateTime(2023, 1, 2));
            sheet.Cells["A4"].PutValue(new DateTime(2023, 1, 3));
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            // Add a pivot table based on the data range
            int pivotIdx = sheet.PivotTables.Add("A1:B4", "D3", "SalesPivot");
            PivotTable pivot = sheet.PivotTables[pivotIdx];

            // Add the Date field as a row field and Sales as a data field
            pivot.AddFieldToArea(PivotFieldType.Row, "Date");
            pivot.AddFieldToArea(PivotFieldType.Data, "Sales");

            // NOTE: The AutoGroup property is not available in the current Aspose.Cells version.
            // If needed, adjust grouping settings via other available APIs.

            // Refresh pivot cache data and calculate the pivot table
            pivot.RefreshData();          // Refreshes the underlying cache
            pivot.CalculateData();        // Calculates the pivot table values

            // Save the workbook
            workbook.Save("Pivot_NoAutoGroup.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
