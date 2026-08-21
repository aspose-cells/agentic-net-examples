// Title: Aspose.Cells .NET: Build a Pivot Table to Summarize Sales by Region
// Description: C# example that creates a new workbook, adds a data sheet with Region and Sales columns, inserts a second worksheet for the pivot, defines the source range (=Data!A1:B9), places the pivot at A3, adds Region as a row field and Sales as a summed data field, shows the result in tabular form, refreshes and calculates the pivot, and saves the file as SalesByRegionPivot.xlsx.
// Keywords: Aspose.Cells pivot table C# | create pivot table .NET | summarize sales by region | add row field Aspose.Cells | add data field Aspose.Cells | pivot table tabular form | RefreshData CalculateData Aspose.Cells | export pivot to Excel
// Common Searches: How to create a pivot table with Aspose.Cells for .NET | Aspose.Cells example grouping sales by region | Add row and data fields to a pivot table using C# | Show pivot table in tabular form Aspose.Cells | Refresh and calculate pivot data programmatically
// Developer Intent: Generate a pivot table that groups sales amounts by region and saves the workbook as an Excel file.
// Use Cases: Produce a sales‑by‑region report for quarterly business reviews. | Create a reusable pivot template that can be refreshed with new data sets. | Export a regional performance summary to Excel for distribution to stakeholders.
// AI Prompts: Write C# code with Aspose.Cells to add a column field for product categories to the existing pivot table and display subtotals. | Update the sample to format the Sales values as currency, apply bold headers, and auto‑fit column widths. | Explain the purpose of RefreshData and CalculateData in Aspose.Cells pivot tables and when each method should be invoked.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;

namespace PivotTableExample
{
    // C# example that creates a new workbook, adds a data sheet with Region and Sales columns, inserts a second worksheet for the pivot, defines the source range (=Data!A1:B9), places the pivot at A3, adds Region as a row field and Sales as a summed data field, shows the result in tabular form, refreshes and calculates the pivot, and saves the file as SalesByRegionPivot.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Prepare source data (Region, Sales)
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Header row
            dataSheet.Cells["A1"].PutValue("Region");
            dataSheet.Cells["B1"].PutValue("Sales");

            // Sample data
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue(1200);
            dataSheet.Cells["A3"].PutValue("South");
            dataSheet.Cells["B3"].PutValue(850);
            dataSheet.Cells["A4"].PutValue("East");
            dataSheet.Cells["B4"].PutValue(950);
            dataSheet.Cells["A5"].PutValue("West");
            dataSheet.Cells["B5"].PutValue(1100);
            dataSheet.Cells["A6"].PutValue("North");
            dataSheet.Cells["B6"].PutValue(700);
            dataSheet.Cells["A7"].PutValue("South");
            dataSheet.Cells["B7"].PutValue(400);
            dataSheet.Cells["A8"].PutValue("East");
            dataSheet.Cells["B8"].PutValue(600);
            dataSheet.Cells["A9"].PutValue("West");
            dataSheet.Cells["B9"].PutValue(500);

            // -------------------------------------------------
            // 2. Add a worksheet to host the pivot table
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // -------------------------------------------------
            // 3. Create the pivot table
            // -------------------------------------------------
            // Data source range: A1:B9 on the Data sheet
            // Destination cell: A3 on the pivot sheet
            int pivotIndex = pivotSheet.PivotTables.Add("=Data!A1:B9", "A3", "SalesByRegion");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // -------------------------------------------------
            // 4. Configure fields:
            //    - Row field: Region
            //    - Data field: Sales (summed)
            // -------------------------------------------------
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: display the pivot table in tabular form for clarity
            pivotTable.ShowInTabularForm();

            // -------------------------------------------------
            // 5. Refresh and calculate the pivot data
            // -------------------------------------------------
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 6. Save the workbook
            // -------------------------------------------------
            workbook.Save("SalesByRegionPivot.xlsx");
        }
    }
}
