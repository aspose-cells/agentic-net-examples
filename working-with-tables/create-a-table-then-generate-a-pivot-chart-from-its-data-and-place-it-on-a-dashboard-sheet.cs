// Title: Create a PivotTable and PivotChart on a Dashboard Sheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to build a workbook, add a data table, generate a PivotTable on a separate sheet, configure rows/columns, insert a column PivotChart linked to the PivotTable on a dashboard sheet, refresh the chart, and save the file as XLSX using Aspose.Cells for C#.
// Keywords: Aspose.Cells | C# | PivotTable | PivotChart | dashboard worksheet | Excel automation | column chart | chart from pivot | save as xlsx | data table
// Common Searches: Aspose.Cells create PivotTable and PivotChart in C# | How to add a dashboard sheet with pivot chart using Aspose.Cells | C# code for PivotChart linked to PivotTable Aspose.Cells | Generate Excel dashboard programmatically Aspose.Cells | Refresh pivot data and chart Aspose.Cells .NET
// Developer Intent: Generate a data table, derive a PivotTable from it, and display the analysis in a PivotChart on a dedicated dashboard worksheet.
// Use Cases: Automated sales reporting: raw data → pivot analysis → visual dashboard in one workbook. | Monthly KPI dashboard that updates when source data changes, without manual Excel interaction. | Self‑service Excel file for business users that combines raw data, pivot calculations, and interactive charts.
// AI Prompts: Write C# code with Aspose.Cells to create a PivotChart that references an existing PivotTable on another worksheet. | Explain how to refresh a PivotTable and its linked PivotChart after modifying source cells using Aspose.Cells. | Show the correct syntax for setting the PivotSource property of a chart to a specific PivotTable name in Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace AsposeCellsPivotChartDemo
{
    // Demonstrates how to build a workbook, add a data table, generate a PivotTable on a separate sheet, configure rows/columns, insert a column PivotChart linked to the PivotTable on a dashboard sheet, refresh the chart, and save the file as XLSX using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 1. Create a data table on the first worksheet
            // -------------------------------------------------
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "Data";

            // Sample data: Category, Product, Sales
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Product");
            dataSheet.Cells["C1"].PutValue("Sales");

            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue("Apple");
            dataSheet.Cells["C2"].PutValue(120);

            dataSheet.Cells["A3"].PutValue("A");
            dataSheet.Cells["B3"].PutValue("Banana");
            dataSheet.Cells["C3"].PutValue(80);

            dataSheet.Cells["A4"].PutValue("B");
            dataSheet.Cells["B4"].PutValue("Carrot");
            dataSheet.Cells["C4"].PutValue(150);

            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue("Date");
            dataSheet.Cells["C5"].PutValue(200);

            // -------------------------------------------------
            // 2. Add a PivotTable on a new worksheet
            // -------------------------------------------------
            Worksheet pivotSheet = workbook.Worksheets.Add("Pivot");
            // Define the source data range (including headers)
            string sourceData = $"=Data!{dataSheet.Cells.MaxDisplayRange.Address}";
            // Add the pivot table (sourceData, destination cell, table name)
            int pivotIndex = pivotSheet.PivotTables.Add(sourceData, "A1", "SalesPivot");
            PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

            // Configure the pivot table: Category as row, Product as column, Sales as data
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Category");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Refresh and calculate the pivot data
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // -------------------------------------------------
            // 3. Create a dashboard sheet and add a PivotChart
            // -------------------------------------------------
            Worksheet dashboardSheet = workbook.Worksheets.Add("Dashboard");
            // Add a column chart (type, upper-left row, upper-left column, lower-right row, lower-right column)
            int chartIndex = dashboardSheet.Charts.Add(ChartType.Column, 1, 0, 20, 10);
            Chart chart = dashboardSheet.Charts[chartIndex];

            // Set the chart's pivot source to the previously created pivot table
            // Format: SheetName!PivotTableName
            chart.PivotSource = $"Pivot!SalesPivot";

            // Refresh the chart to pull data from the pivot table
            chart.RefreshPivotData();

            // -------------------------------------------------
            // 4. Save the workbook
            // -------------------------------------------------
            workbook.Save("PivotChartDashboard.xlsx", SaveFormat.Xlsx);
        }
    }
}
