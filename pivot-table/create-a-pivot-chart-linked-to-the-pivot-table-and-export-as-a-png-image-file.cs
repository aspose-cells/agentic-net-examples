// Title: Export a Pivot‑Chart Linked to a Pivot Table as PNG with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample sales data, build a pivot table, place a column chart linked to that pivot, refresh the chart data, and save the chart as a PNG image while also saving the workbook as an XLSX file using Aspose.Cells for C#.
// Keywords: Aspose.Cells pivot chart | export chart to PNG | pivot table chart C# | Aspose.Cells .NET image export | programmatic chart generation | pivot chart automation | C# Excel chart export
// Common Searches: Aspose.Cells export pivot chart as PNG | Create pivot chart from pivot table C# | Save Excel chart image with Aspose.Cells | Link chart to pivot table programmatically | Generate PNG chart from workbook using Aspose
// Developer Intent: Create a pivot chart bound to a pivot table and output it as a PNG file.
// Use Cases: Add a sales‑summary chart image to automated PDF reports. | Produce thumbnail charts for a web dashboard that updates from Excel data. | Attach chart images to email alerts without opening Excel.
// AI Prompts: Write C# code with Aspose.Cells that builds a pivot table, adds a column chart linked to it, and saves the chart as a PNG file. | Explain the role of Chart.PivotSource and Chart.RefreshPivotData when connecting a chart to a pivot table in Aspose.Cells. | Show how to set custom width, height, and DPI for a PNG export of a pivot chart using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPivotChartExport
{
    // Demonstrates how to create a workbook, add sample sales data, build a pivot table, place a column chart linked to that pivot, refresh the chart data, and save the chart as a PNG image while also saving the workbook as an XLSX file using Aspose.Cells for C#.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet (will hold the source data)
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data for the pivot table
                dataSheet.Cells["A1"].PutValue("Category");
                dataSheet.Cells["B1"].PutValue("Sales");
                dataSheet.Cells["A2"].PutValue("Fruit");
                dataSheet.Cells["B2"].PutValue(1200);
                dataSheet.Cells["A3"].PutValue("Vegetable");
                dataSheet.Cells["B3"].PutValue(800);
                dataSheet.Cells["A4"].PutValue("Fruit");
                dataSheet.Cells["B4"].PutValue(1500);
                dataSheet.Cells["A5"].PutValue("Vegetable");
                dataSheet.Cells["B5"].PutValue(600);

                // Add a new worksheet that will contain the pivot table and the chart
                Worksheet pivotSheet = workbook.Worksheets.Add("PivotAndChart");

                // Create a pivot table based on the data range A1:B5
                // The pivot table will be placed starting at cell D1 in the pivot sheet
                int pivotIndex = pivotSheet.PivotTables.Add("Data!A1:B5", "D1", "PivotTable1");
                PivotTable pivotTable = pivotSheet.PivotTables[pivotIndex];

                // Configure the pivot table: Category on rows, Sales on data
                pivotTable.AddFieldToArea(PivotFieldType.Row, 0);   // Column 0 -> Category
                pivotTable.AddFieldToArea(PivotFieldType.Data, 1); // Column 1 -> Sales

                // Refresh and calculate the pivot table data
                pivotTable.RefreshData();
                pivotTable.CalculateData();

                // Add a chart to the same worksheet (positioned below the pivot table)
                int chartIndex = pivotSheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
                Chart chart = pivotSheet.Charts[chartIndex];

                // Link the chart to the pivot table
                chart.PivotSource = "PivotTable1";

                // Ensure the chart picks up the latest pivot data
                chart.RefreshPivotData();

                // Export the chart as a PNG image file (default format is PNG)
                chart.ToImage("PivotChart.png");

                // Optionally, save the workbook to verify the result
                workbook.Save("PivotChartDemo.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
