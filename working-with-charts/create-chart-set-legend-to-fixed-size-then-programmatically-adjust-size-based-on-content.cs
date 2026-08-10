// Title: Aspose.Cells for .NET: Create a Column Chart with Fixed Legend Size and Dynamically Adjust Width & Height
// Description: This example shows how to build a column chart in an Excel workbook using Aspose.Cells, disable the automatic legend sizing, set an initial width and height, then programmatically enlarge the legend based on the longest series name and the number of series before saving the file.
// Keywords: Aspose.Cells | .NET | C# | chart legend size | fixed legend size | dynamic legend resizing | column chart | Excel chart legend width | Excel chart legend height | disable automatic legend sizing | programmatic legend adjustment
// Common Searches: Aspose.Cells change legend width C# | Resize Excel chart legend programmatically .NET | Set fixed legend size then auto expand Aspose.Cells | Calculate legend height based on series count Aspose.Cells | Disable automatic legend sizing in Aspose.Cells chart
// Developer Intent: The developer wants to generate a column chart and ensure the legend automatically expands to accommodate the longest series label and the total number of series.
// Use Cases: Create a chart with a preset legend size and automatically widen it when a series name exceeds the initial width. | Increase legend height on the fly when the chart contains more series than the original legend can display. | Recalculate the chart after resizing the legend so the final workbook reflects the updated dimensions.
// AI Prompts: Write C# code with Aspose.Cells that adds a column chart, disables automatic legend sizing, and expands the legend width based on the longest series name. | Provide an Aspose.Cells example that computes the required legend height from the series count and updates both width and height before saving the workbook. | Explain how to set a fixed legend size and then programmatically adjust its dimensions in an Excel chart using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This example shows how to build a column chart in an Excel workbook using Aspose.Cells, disable the automatic legend sizing, set an initial width and height, then programmatically enlarge the legend based on the longest series name and the number of series before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Alpha");
            sheet.Cells["A3"].PutValue("Beta");
            sheet.Cells["A4"].PutValue("Gamma");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Set legend to a fixed size (disable automatic sizing)
            chart.Legend.IsAutomaticSize = false;
            chart.Legend.Position = LegendPositionType.Right;
            chart.Legend.WidthPixel = 150;   // initial width
            chart.Legend.HeightPixel = 100;  // initial height

            // ---- Programmatically adjust legend size based on its content ----

            // Determine the longest legend entry text (use series names)
            int maxTextLength = 0;
            foreach (Series series in chart.NSeries)
            {
                string name = series.Name ?? string.Empty;
                if (name.Length > maxTextLength)
                    maxTextLength = name.Length;
            }

            // Approximate required width (7 pixels per character) + padding
            int requiredWidth = maxTextLength * 7 + 20;
            if (requiredWidth > chart.Legend.WidthPixel)
                chart.Legend.WidthPixel = requiredWidth;

            // Approximate required height (20 pixels per entry) + padding
            int requiredHeight = chart.NSeries.Count * 20 + 20;
            if (requiredHeight > chart.Legend.HeightPixel)
                chart.Legend.HeightPixel = requiredHeight;

            // Recalculate the chart to apply layout changes
            chart.Calculate();

            // Save the workbook
            workbook.Save("ChartLegendDynamicSize.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
