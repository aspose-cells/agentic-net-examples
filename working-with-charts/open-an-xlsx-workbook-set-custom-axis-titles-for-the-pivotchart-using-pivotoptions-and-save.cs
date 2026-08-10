// Title: Aspose.Cells for .NET: Create a PivotChart with custom axis titles and enable drop zones via PivotOptions (C#)
// Description: This example demonstrates how to generate an XLSX workbook, add source data, create a pivot table, insert a column PivotChart, enable drop zones using PivotOptions, set custom category and value axis titles, refresh the chart data, and save the file as PivotChartWithCustomAxis.xlsx using Aspose.Cells for .NET (C#).
// Keywords: Aspose.Cells | C# | PivotChart | custom axis title | PivotOptions | drop zones | refresh pivot chart | create pivot table | column chart | save XLSX | Aspose.Cells API | chart axis customization
// Common Searches: Aspose.Cells set pivot chart axis title C# | Enable drop zones on PivotChart Aspose.Cells | Refresh pivot chart after changing PivotOptions | How to add custom category axis text in Aspose.Cells | Create pivot chart from worksheet using Aspose.Cells .NET
// Developer Intent: Create and save an XLSX workbook that contains a pivot chart with custom axis titles and visible drop zones using Aspose.Cells for .NET.
// Use Cases: Automated financial reporting that requires pivot charts with clear, custom axis labels for stakeholders. | Dynamic dashboards where users can rearrange fields via drop zones while preserving customized axis titles. | Batch generation of sales analysis workbooks that need refreshed pivot chart data after source updates.
// AI Prompts: Write C# code with Aspose.Cells to build a workbook, add a pivot table, create a column PivotChart, enable drop zones via PivotOptions, set custom category and value axis titles, refresh the chart, and save the file. | Explain step‑by‑step how PivotOptions affect interactivity of a PivotChart and how to set custom axis titles in Aspose.Cells for .NET. | Provide a concise guide for developers to add a pivot table, link it to a chart, customize axis titles, enable drop zones, refresh the chart data, and export to XLSX using Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// This example demonstrates how to generate an XLSX workbook, add source data, create a pivot table, insert a column PivotChart, enable drop zones using PivotOptions, set custom category and value axis titles, refresh the chart data, and save the file as PivotChartWithCustomAxis.xlsx using Aspose.Cells for .NET (C#).
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate source data for the pivot table
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["B2"].PutValue(10);
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["B3"].PutValue(20);
        worksheet.Cells["A4"].PutValue("A");
        worksheet.Cells["B4"].PutValue(30);
        worksheet.Cells["A5"].PutValue("B");
        worksheet.Cells["B5"].PutValue(40);

        // Add a pivot table based on the data range
        int pivotIndex = worksheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
        PivotTable pivotTable = worksheet.PivotTables[pivotIndex];
        pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
        pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

        // Add a chart and link it to the pivot table
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];
        chart.PivotSource = "PivotTable1";

        // Use PivotOptions (e.g., enable drop zones)
        PivotOptions pivotOptions = chart.PivotOptions;
        pivotOptions.DropZonesVisible = true;

        // Set custom axis titles for the pivot chart
        chart.CategoryAxis.Title.Text = "Custom Category Axis";
        chart.ValueAxis.Title.Text = "Custom Value Axis";

        // Refresh chart data from the pivot table
        chart.RefreshPivotData();

        // Save the workbook
        workbook.Save("PivotChartWithCustomAxis.xlsx");
    }
}
