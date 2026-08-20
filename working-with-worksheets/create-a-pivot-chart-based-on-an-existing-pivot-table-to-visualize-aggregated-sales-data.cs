// Title: Create a PivotChart from a PivotTable with Aspose.Cells for .NET (C#)
// Description: C# code that builds a workbook, inserts sample sales rows, creates a PivotTable (Region rows, Product columns, Sum of Sales), adds a linked column PivotChart, enables interactive drop‑zone controls, refreshes the chart data, and saves the file as PivotChartDemo.xlsx using Aspose.Cells.
// Keywords: Aspose.Cells pivot chart C# | create pivot chart Aspose.Cells | link pivot chart to pivot table .NET | pivot chart drop zones | refresh pivot chart Aspose | C# Excel pivot table example | sales data pivot chart | Aspose.Cells chart types | Excel automation Aspose.Cells
// Common Searches: Aspose.Cells how to add a PivotChart to a workbook | C# create PivotTable and PivotChart with Aspose.Cells | Enable drop zone controls on PivotChart Aspose | Refresh PivotChart after changing PivotTable data Aspose.Cells | Example of sales data pivot chart using Aspose.Cells .NET
// Developer Intent: Generate a column PivotChart that automatically reflects the aggregated sales values defined in a PivotTable.
// Use Cases: Produce an executive‑ready column chart showing regional product sales derived from a PivotTable. | Create an interactive Excel file where users can modify categories, series, and filters directly on the chart via drop‑zone controls. | Export a single workbook containing both a PivotTable and its linked chart for distribution to stakeholders.
// AI Prompts: Write C# code with Aspose.Cells to add a PivotChart linked to an existing PivotTable and enable all drop‑zone options. | Show how to refresh a PivotChart after updating the underlying PivotTable data using Aspose.Cells for .NET. | Explain how to change the chart type and customize PivotOptions for a PivotChart created from sales data.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartDemo
{
    // C# code that builds a workbook, inserts sample sales rows, creates a PivotTable (Region rows, Product columns, Sum of Sales), adds a linked column PivotChart, enables interactive drop‑zone controls, refreshes the chart data, and saves the file as PivotChartDemo.xlsx using Aspose.Cells.
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

            // Sample rows
            dataSheet.Cells["A2"].PutValue("North");
            dataSheet.Cells["B2"].PutValue("Laptop");
            dataSheet.Cells["C2"].PutValue(1200);

            dataSheet.Cells["A3"].PutValue("North");
            dataSheet.Cells["B3"].PutValue("Phone");
            dataSheet.Cells["C3"].PutValue(800);

            dataSheet.Cells["A4"].PutValue("South");
            dataSheet.Cells["B4"].PutValue("Laptop");
            dataSheet.Cells["C4"].PutValue(1500);

            dataSheet.Cells["A5"].PutValue("South");
            dataSheet.Cells["B5"].PutValue("Phone");
            dataSheet.Cells["C5"].PutValue(900);

            // -------------------------------------------------
            // 2. Create a PivotTable on a new worksheet
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("PivotTable");

            // Define the source data range (including headers)
            string sourceData = "=SalesData!A1:C5";

            // Add the pivot table; top‑left corner at cell A3
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A3", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot fields:
            // Row axis -> Region
            // Column axis -> Product
            // Data area -> Sum of Sales
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Optional: calculate the pivot data now
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 3. Add a PivotChart linked to the PivotTable
            // -------------------------------------------------
            // Add a column chart positioned on the same pivot sheet
            int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 15, 0, 30, 15);
            Chart chart = pivotSheet.Charts[chartIndex];

            // Link the chart to the pivot table
            chart.PivotSource = "PivotTable!SalesPivot";

            // Refresh the chart so it pulls data from the pivot table
            chart.RefreshPivotData();

            // Optional: enable pivot controls on the chart
            PivotOptions pivotOptions = chart.PivotOptions;
            pivotOptions.DropZonesVisible = true;
            pivotOptions.DropZoneFilter = true;
            pivotOptions.DropZoneCategories = true;
            pivotOptions.DropZoneSeries = true;
            pivotOptions.DropZoneData = true;

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotChartDemo.xlsx");
        }
    }
}
