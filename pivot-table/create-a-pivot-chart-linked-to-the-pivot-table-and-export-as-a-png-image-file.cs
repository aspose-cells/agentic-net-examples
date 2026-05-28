using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsPivotChartExport
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            dataSheet.Cells["A1"].PutValue("Category");
            dataSheet.Cells["B1"].PutValue("Value");
            dataSheet.Cells["A2"].PutValue("A");
            dataSheet.Cells["B2"].PutValue(10);
            dataSheet.Cells["A3"].PutValue("B");
            dataSheet.Cells["B3"].PutValue(20);
            dataSheet.Cells["A4"].PutValue("A");
            dataSheet.Cells["B4"].PutValue(30);
            dataSheet.Cells["A5"].PutValue("B");
            dataSheet.Cells["B5"].PutValue(40);

            // Add a pivot table based on the data range
            int pivotIndex = dataSheet.PivotTables.Add("A1:B5", "D1", "PivotTable1");
            PivotTable pivotTable = dataSheet.PivotTables[pivotIndex];
            pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Category as row field
            pivotTable.AddFieldToArea(PivotFieldType.Data, 1);  // Value as data field

            // Add a chart that will be linked to the pivot table
            int chartIndex = dataSheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = dataSheet.Charts[chartIndex];

            // Set the pivot source for the chart (makes it a pivot chart)
            chart.PivotSource = "PivotTable1";

            // Refresh the chart data from the pivot table
            chart.RefreshPivotData();

            // Export the pivot chart to a PNG image file
            chart.ToImage("PivotChart.png", ImageType.Png);

            // Optionally save the workbook to verify the chart and pivot table are stored
            workbook.Save("PivotChartDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Pivot chart exported to PivotChart.png and workbook saved.");
        }
    }
}