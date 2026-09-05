// Title: Create a column chart in Aspose.Cells for .NET and programmatically resize the legend based on the longest series name
// AI Prompts: Generate C# code with Aspose.Cells that adds a column chart, disables automatic legend sizing, and sets Legend.WidthPixel and HeightPixel according to the length of the longest series name. | Show how to estimate pixel width from a text string and apply that calculation to adjust the legend dimensions of an Excel chart using Aspose.Cells.
// Common Searches: how to set a fixed legend size and then resize it dynamically with Aspose.Cells C# | Aspose.Cells calculate legend width from series name length in a column chart | C# example for adjusting Excel chart legend pixel dimensions using Aspose.Cells
// Tags: Aspose.Cells column chart legend pixel sizing | disable automatic legend size Aspose.Cells | calculate legend width from series name C# | dynamic legend resizing Aspose.Cells chart | Excel chart legend customization Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLegendResizeDemo
{
    // The example creates a workbook, fills it with sample data, adds a column chart, disables automatic legend sizing, computes the longest series name, and then updates the legend's WidthPixel and HeightPixel to fit the text before saving the file as ChartLegendDynamicSize.xlsx.
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
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.SetChartDataRange("A1:B4", true);

                // Configure the legend with a fixed size first
                Legend legend = chart.Legend;
                legend.IsAutomaticSize = false;          // disable automatic sizing
                legend.Position = LegendPositionType.Bottom;
                legend.WidthPixel = 200;                  // initial width (pixels)
                legend.HeightPixel = 50;                  // initial height (pixels)

                // Adjust legend size based on the longest series name
                int maxTextLength = 0;
                foreach (Series series in chart.NSeries)
                {
                    int length = series.Name?.Length ?? 0;
                    if (length > maxTextLength)
                        maxTextLength = length;
                }

                // If there are entries, resize the legend
                if (maxTextLength > 0)
                {
                    const int approxCharWidth = 7; // rough average pixel width per character
                    legend.WidthPixel = maxTextLength * approxCharWidth + 20; // add padding
                    legend.HeightPixel = 30; // enough height for a single line
                }

                // Ensure the chart layout reflects the changes
                chart.Calculate();

                // Save the workbook
                workbook.Save("ChartLegendDynamicSize.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
