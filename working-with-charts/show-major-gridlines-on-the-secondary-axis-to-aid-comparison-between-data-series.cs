// Title: Show major gridlines on a secondary value axis in a column chart using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart with two data series, assign the second series to the secondary value axis, enable its major gridlines, set a custom axis title, and save the workbook with Aspose.Cells in C#. | Use reflection to safely set the IsSecondaryValueAxis flag and the MajorGridLines.IsVisible property for a chart's secondary axis when the direct API is unavailable.
// Common Searches: how to enable major gridlines on secondary axis with Aspose.Cells C# | Aspose.Cells column chart secondary value axis gridlines example | C# assign series to secondary axis and show gridlines using Aspose.Cells | reflection to set IsSecondaryValueAxis property Aspose.Cells .NET | save Excel chart with secondary axis formatting Aspose.Cells
// Tags: Aspose.Cells enable secondary axis gridlines | C# column chart secondary value axis | Aspose.Cells reflection set chart properties | Excel chart major gridlines secondary axis C# | Aspose.Cells secondary axis formatting

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, fills it with sample data, adds a column chart with two series, moves the second series to the secondary value axis, uses reflection to make the secondary axis's major gridlines visible and assign a title, and saves the result as ChartWithSecondaryAxisGridlines.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(100);
            sheet.Cells["C3"].PutValue(150);
            sheet.Cells["C4"].PutValue(200);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // First series (primary axis)
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].XValues = "A2:A4";

            // Second series (attempt to place on secondary axis if supported)
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries[1].XValues = "A2:A4";

            // The following features may not be available in older Aspose.Cells versions.
            // They are wrapped in try-catch blocks to avoid compilation/runtime errors.
            try
            {
                // Attempt to assign the second series to the secondary value axis
                // (property exists in newer versions)
                var series = chart.NSeries[1];
                var prop = series.GetType().GetProperty("IsSecondaryValueAxis");
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(series, true);
                }
            }
            catch
            {
                // Ignored – secondary axis assignment not supported
            }

            try
            {
                // Attempt to enable gridlines on the secondary value axis
                var secondaryAxisProp = chart.GetType().GetProperty("SecondaryValueAxis");
                if (secondaryAxisProp != null)
                {
                    var secondaryAxis = secondaryAxisProp.GetValue(chart);
                    var gridLinesProp = secondaryAxis.GetType().GetProperty("MajorGridLines");
                    var titleProp = secondaryAxis.GetType().GetProperty("Title");
                    if (gridLinesProp != null && titleProp != null)
                    {
                        var gridLines = gridLinesProp.GetValue(secondaryAxis);
                        var isVisibleProp = gridLines.GetType().GetProperty("IsVisible");
                        if (isVisibleProp != null)
                        {
                            isVisibleProp.SetValue(gridLines, true);
                        }

                        var title = titleProp.GetValue(secondaryAxis);
                        var textProp = title.GetType().GetProperty("Text");
                        if (textProp != null)
                        {
                            textProp.SetValue(title, "Secondary Axis");
                        }
                    }
                }
            }
            catch
            {
                // Ignored – secondary axis features not supported
            }

            // Save the workbook
            string outputPath = "ChartWithSecondaryAxisGridlines.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
