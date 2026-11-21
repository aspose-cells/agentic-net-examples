// Title: Set a fixed 0‑to‑1000 Y‑axis range on a column chart using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set chart value axis minimum and maximum | how to disable automatic Y axis scaling in a bar chart with Aspose.Cells | C# code to define custom Y axis limits for a column chart in Excel | Aspose.Cells chart axis range 0 to 1000 example | programmatically adjust value axis range in Aspose.Cells .NET
// Tags: custom Y axis limits Aspose.Cells | column chart axis scaling .NET | disable automatic value axis Aspose.Cells | Excel chart axis range configuration C# | fixed 0‑1000 Y axis Aspose.Cells chart

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates sample data, adds a vertical column chart, disables automatic scaling on the value axis, sets the Y‑axis minimum to 0 and maximum to 1000, and saves the file as BarChartWithCustomYAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the bar chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(200);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(500);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(800);

            // Add a column chart (vertical bar chart) to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series and categories for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---- Adjust Y‑axis (value axis) scale ----
            // Disable automatic scaling for the value axis
            chart.ValueAxis.IsAutomaticMinValue = false;
            chart.ValueAxis.IsAutomaticMaxValue = false;
            // Set the desired minimum and maximum values
            chart.ValueAxis.MinValue = 0;
            chart.ValueAxis.MaxValue = 1000;
            // -----------------------------------------

            // Save the workbook with the customized chart
            workbook.Save("BarChartWithCustomYAxis.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
