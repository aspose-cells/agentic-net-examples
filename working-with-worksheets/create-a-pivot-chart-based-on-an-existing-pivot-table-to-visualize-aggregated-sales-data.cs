// Title: Aspose.Cells for .NET: Create a PivotChart from a PivotTable to Visualize Sales Data (C#)
// Description: This example creates a workbook, adds sample sales rows, builds a PivotTable with Region rows, Product columns and summed Sales, then generates a column PivotChart linked to the table, configures drop‑zone options, refreshes the data, and saves the file as PivotChartExample.xlsx.
// Keywords: Aspose.Cells C# pivot chart | Aspose.Cells pivot table example | create pivot chart .NET | visualize sales data Aspose | programmatic Excel chart | pivot chart drop zones | refresh pivot data | column chart from pivot table | Excel automation C# | Aspose.Cells tutorial
// Common Searches: Aspose.Cells create pivot chart from pivot table C# | How to link a PivotChart to a PivotTable using Aspose.Cells | Set drop zones on a PivotChart with Aspose.Cells | Refresh PivotChart data programmatically Aspose.Cells | Example of sales dashboard with PivotTable and PivotChart in .NET
// Developer Intent: Generate a PivotChart that is bound to a PivotTable for displaying aggregated sales figures.
// Use Cases: Automate a sales dashboard workbook that updates charts when source data changes. | Export a refreshed PivotChart as an image for inclusion in reports or presentations. | Provide end‑users with interactive drop‑zone controls to reshape the chart directly in Excel.
// AI Prompts: Show how to change the PivotChart type to a line chart and set custom series colors in Aspose.Cells for .NET. | Provide code to save the refreshed PivotChart as a PNG file after calling RefreshPivotData. | Explain how to configure the workbook so the PivotChart automatically refreshes on file open.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartExample
{
    // This example creates a workbook, adds sample sales rows, builds a PivotTable with Region rows, Product columns and summed Sales, then generates a column PivotChart linked to the table, configures drop‑zone options, refreshes the data, and saves the file as PivotChartExample.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Add sample sales data to the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "SalesData";

            // Header row
            dataSheet.Cells["A1"].PutValue("Region");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            // Sample data
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("North");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("South");
            dataSheet.Cells["B4"].PutValue("Apple");
            dataSheet.Cells["C4"].PutValue(1500);

            dataSheet.Cells["A5"].PutValue("South");
            dataSheet.Cells["B5"].PutValue("Banana");
            dataSheet.Cells["C5"].PutValue(700);

            // -------------------------------------------------
            // 2. Create a PivotTable on a new worksheet
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            // Define the source data range (including headers)
            string sourceData = "=SalesData!A1:C5";
            // Add the pivot table; top‑left cell of the report will be A1
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields:
            // Row axis -> Region
            // Column axis -> Product
            // Data area -> Sum of Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: set the layout to compact form
            pivotTable.ShowInCompactForm();

            // -------------------------------------------------
            // 3. Add a PivotChart that uses the created PivotTable
            // -------------------------------------------------
            // Add a column chart positioned on the same worksheet
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Link the chart to the pivot table
            chart.PivotSource = "Pivot!SalesPivot";

            // Refresh the chart so it pulls data from the pivot table
            chart.RefreshPivotData();

            // (Optional) Enable pivot controls on the chart
            PivotOptions options = chart.PivotOptions;
            options.DropZonesVisible = true;
            options.DropZoneCategories = true;
            options.DropZoneSeries = true;
            options.DropZoneData = true;
            options.DropZoneFilter = true;

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotChartExample.xlsx");
        }
    }
}
