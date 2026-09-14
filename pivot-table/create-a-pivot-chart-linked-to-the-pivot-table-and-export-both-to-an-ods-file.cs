// Title: Create a pivot table with a linked column chart and export both to an ODS file using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that builds a pivot table from a data range, adds a column chart linked via Chart.PivotSource, refreshes the chart data, and saves the workbook as an ODS file while preserving the pivot information. | Change the chart type to a line chart in the existing Aspose.Cells example, keep the pivot‑chart linkage, and export the workbook to ODS. | Add a second numeric field to the pivot table, update the chart series to reflect the new field, and save the workbook to ODS with all pivot data included.
// Common Searches: how to link a pivot chart to a pivot table using Aspose.Cells C# | export pivot table and chart to ODS format with Aspose.Cells | Aspose.Cells OdsSaveOptions include pivot tables example | create column chart from pivot table programmatically Aspose.Cells | C# sample for saving workbook with pivot chart to ODS
// Tags: Aspose.Cells pivot table creation C# | Chart.PivotSource linking pivot chart | export workbook to ODS with pivot data | OdsSaveOptions include pivot tables | column chart from pivot table Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Ods;

// The example demonstrates how to generate a workbook, populate it with sample data, create a pivot table, add a column chart linked to that pivot via the Chart.PivotSource property, refresh the chart, configure OdsSaveOptions to retain pivot tables, and finally save the workbook as an ODS file.
class PivotChartToOds
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Name = "Data";

        // Populate sample data for the pivot table
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("A");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["A5"].PutValue("B");
        sheet.Cells["B5"].PutValue(40);

        // Add a pivot table based on the data range
        int pivotIndex = sheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = sheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Add a column chart and link it to the pivot table
        int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
        Chart chart = sheet.Charts[chartIndex];
        chart.PivotSource = "PivotTable1";          // Link chart to the pivot table (Chart.PivotSource rule)
        chart.RefreshPivotData();                  // Refresh chart data from the pivot table (Chart.RefreshPivotData rule)

        // Save the workbook as ODS, ensuring pivot tables are included
        OdsSaveOptions saveOptions = new OdsSaveOptions(); // OdsSaveOptions constructor rule
        saveOptions.IgnorePivotTables = false;              // Include pivot tables in the ODS file
        workbook.Save("PivotChartDemo.ods", saveOptions);  // Save using the provided save rule
    }
}
