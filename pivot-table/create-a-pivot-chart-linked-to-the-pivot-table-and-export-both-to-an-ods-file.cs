using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Ods;

namespace AsposeCellsPivotChartExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table (A1:B5)
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1200);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(850);
            sheet.Cells["A4"].PutValue("Orange");
            sheet.Cells["B4"].PutValue(950);
            sheet.Cells["A5"].PutValue("Apple");
            sheet.Cells["B5"].PutValue(300);

            // Add a pivot table based on the data range and place it at D1
            int pivotIndex = sheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure the pivot table: Product as row field, Sales as data field
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Product
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Column 1 -> Sales

            // Refresh and calculate the pivot data so the chart can use it
            pivotTable.RefreshData();
            pivotTable.CalculateData();

            // Add a column chart that will be linked to the pivot table
            int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Link the chart to the pivot table (creates a PivotChart)
            chart.PivotSource = "PivotTable1";

            // Refresh chart data from the pivot table
            chart.RefreshPivotData();

            // Save the workbook as ODS, ensuring pivot tables are included
            OdsSaveOptions saveOptions = new OdsSaveOptions();
            saveOptions.IgnorePivotTables = false; // include pivot tables and pivot chart
            workbook.Save("PivotChartOutput.ods", saveOptions);
        }
    }
}