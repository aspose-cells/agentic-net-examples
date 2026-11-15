// Title: Create a reusable column chart template with right-aligned legend and blue data-label font using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set chart legend position to right | Add data labels with custom font color and size to a column chart in Aspose.Cells | Generate a chart template with predefined legend and data label settings using Aspose.Cells for .NET | How to customize data label appearance in Aspose.Cells chart programmatically
// Tags: set chart legend position right Aspose.Cells | customize data label font Aspose.Cells C# | column chart template generation Aspose.Cells | enable data labels show value Aspose.Cells | predefined chart template with legend Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, adds a column chart with sample data, positions the legend on the right side, enables data labels to display each point's value, customizes the label font to blue size 10, and saves the workbook as WorkbookWithChart.xlsx.
class ChartTemplateCreator
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position the legend on the right side
            chart.Legend.Position = LegendPositionType.Right;

            // Enable data labels and show values
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true; // Show the value of each point

            // Optional: customize data label font
            series.DataLabels.Font.Color = Color.Blue;
            series.DataLabels.Font.Size = 10;

            // Save the workbook (the chart is stored inside)
            workbook.Save("WorkbookWithChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
