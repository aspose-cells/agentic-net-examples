// Title: Create a Column‑Line Combo Chart with Aspose.Cells for .NET
// Description: This example builds an Excel workbook, fills it with month, sales and profit data, adds a column chart, converts the profit series to a line, customizes its markers, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells combo chart | column and line series .NET | change series type Aspose.Cells | custom line markers Excel chart | Aspose.Cells chart API example | Excel combo chart code | Aspose.Cells .NET tutorial
// Common Searches: Aspose.Cells create combo chart column line | how to change series type to line in Aspose.Cells | add markers to line series Aspose.Cells chart | combo chart example Aspose.Cells for .NET | Aspose.Cells column chart with secondary line series
// Developer Intent: Generate an Excel file that displays sales as columns and profit as a line on the same chart.
// Use Cases: Business dashboards that compare sales volume (columns) with profit trend (line) in a single view. | Reports requiring two metrics with different scales without using separate charts. | Highlighting profit fluctuations by applying custom markers while keeping sales data in column format.
// AI Prompts: Show how to add a secondary Y‑axis for the line series in this combo chart. | Provide code that sets distinct colors for the column and line series and adds a chart title. | Explain how to export the generated combo chart as a PNG image using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example builds an Excel workbook, fills it with month, sales and profit data, adds a column chart, converts the profit series to a line, customizes its markers, and saves the file as ComboChart.xlsx using Aspose.Cells for .NET.
class ComboChartDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate header row
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["C1"].PutValue("Profit");

            // Sample data
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] sales = { 100, 150, 130, 170, 160 };
            double[] profit = { 20, 30, 25, 35, 28 };

            // Fill worksheet with data
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 2, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 2, 1].PutValue(sales[i]);   // Column B
                sheet.Cells[i + 2, 2].PutValue(profit[i]);  // Column C
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A6";

            // Add Sales series (column)
            chart.NSeries.Add("B2:B6", true);

            // Add Profit series (will be changed to line)
            chart.NSeries.Add("C2:C6", true);

            // Convert the second series to a line type for a combo chart
            chart.NSeries[1].Type = ChartType.Line;

            // Customize the line series markers
            chart.NSeries[1].Marker.MarkerStyle = ChartMarkerType.Circle;
            // Size property may not be available in some versions; omitted for compatibility
            chart.NSeries[1].Marker.ForegroundColor = Color.Blue;

            // Save the workbook
            string outputPath = "ComboChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
