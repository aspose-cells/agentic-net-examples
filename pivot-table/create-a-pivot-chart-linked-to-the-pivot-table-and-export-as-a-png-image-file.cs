// Title: Export a Pivot Chart Linked to a Pivot Table as PNG with Aspose.Cells for .NET
// Description: Demonstrates how to create a workbook, add sample data, build a pivot table, insert a column chart linked to that pivot table, refresh the chart data, and export the chart directly to a PNG file using Aspose.Cells for C#.
// Keywords: Aspose.Cells pivot chart PNG | C# export pivot chart image | link chart to pivot table Aspose | Aspose.Cells chart refresh | pivot chart to image .NET | Aspose.Cells example PNG export
// Common Searches: Aspose.Cells export pivot chart as PNG C# | how to link chart to pivot table Aspose.Cells | C# create pivot table and chart then save as image | Aspose.Cells generate PNG from pivot chart | export Excel pivot chart to PNG programmatically
// Developer Intent: Create a pivot chart bound to a pivot table and save the chart as a PNG image file using Aspose.Cells for .NET.
// Use Cases: Automate generation of sales‑summary charts and embed PNGs in PowerPoint decks. | Produce daily dashboard snapshots as PNG images for email or web reporting. | Create printable chart graphics from server‑side Excel data without installing Excel.
// AI Prompts: Write C# code with Aspose.Cells that builds a pivot table, adds a column chart linked to it, refreshes the data, and exports the chart as a PNG file. | Explain step‑by‑step how to attach a pie chart to an existing pivot table in Aspose.Cells and save the result as a high‑resolution PNG. | Show how to iterate over all pivot tables in a workbook, generate corresponding charts, and batch export each chart to separate PNG files using Aspose.Cells.

using System;
using System.Drawing.Imaging;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Charts;

// Demonstrates how to create a workbook, add sample data, build a pivot table, insert a column chart linked to that pivot table, refresh the chart data, and export the chart directly to a PNG file using Aspose.Cells for C#.
class PivotChartToPng
{
    static void Main()
    {
        try
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

            // Link the chart to the pivot table
            chart.PivotSource = "PivotTable1";

            // Refresh the chart so it pulls data from the pivot table
            chart.RefreshPivotData();

            // Export the pivot chart to a PNG image file (default format is PNG)
            chart.ToImage("PivotChart.png");

            // Optionally save the workbook (not required for the image export)
            workbook.Save("PivotChartDemo.xlsx", SaveFormat.Xlsx);

            Console.WriteLine("Pivot chart exported to PivotChart.png successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
