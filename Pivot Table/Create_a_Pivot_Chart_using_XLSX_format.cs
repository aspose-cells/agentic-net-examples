using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

namespace PivotChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

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
            int pivotIdx = sheet.PivotTables.Add("=A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIdx];
            // Place "Category" in rows and "Value" in data area
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);

            // Add a column chart that will be linked to the pivot table
            int chartIdx = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the chart's source to the pivot table (makes it a PivotChart)
            chart.PivotSource = "PivotTable1";

            // Optional: configure pivot controls displayed on the chart
            PivotOptions options = chart.PivotOptions;
            options.DropZonesVisible = true;
            options.DropZoneCategories = true;
            options.DropZoneData = true;
            options.DropZoneFilter = true;
            options.DropZoneSeries = true;

            // Refresh chart data from the pivot table
            chart.RefreshPivotData();

            // Save the workbook in XLSX format
            workbook.Save("PivotChartExample.xlsx", SaveFormat.Xlsx);
        }
    }
}