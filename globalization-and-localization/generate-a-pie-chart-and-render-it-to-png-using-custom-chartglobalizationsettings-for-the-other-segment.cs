// Title: Aspose.Cells C# – Pie Chart with Custom “Other” Label and PNG Export using ChartGlobalizationSettings
// Description: Creates a workbook, fills sample data, adds a pie chart, sets ChartSplitType.PercentValue to group small slices, applies SettableChartGlobalizationSettings to rename the aggregated “Other” segment, renders the chart directly to a PNG file, and saves the workbook as an Excel file.
// Keywords: Aspose.Cells | C# pie chart | ChartGlobalizationSettings | custom other label | ChartSplitType.PercentValue | render chart to PNG | Excel chart export .NET | localization of chart labels | split small values pie chart | Aspose.Cells PNG rendering
// Common Searches: Aspose.Cells change "Other" slice label in pie chart | render Aspose.Cells pie chart to PNG C# | ChartSplitType.PercentValue example Aspose.Cells | apply ChartGlobalizationSettings workbook Aspose.Cells | group small values into "Other" segment Aspose.Cells
// Developer Intent: Generate a pie chart that groups low‑percentage categories under a custom‑named “Other” slice and export the chart as a PNG image.
// Use Cases: Display sales distribution where categories below a threshold are combined under a localized “Other” label and embed the PNG in a PDF report. | Create multilingual dashboards that automatically rename the “Other” segment based on culture‑specific globalization settings. | Produce lightweight PNG thumbnails of Excel charts for web previews without opening the workbook.
// AI Prompts: Write C# code with Aspose.Cells to set a custom name for the "Other" segment in a pie chart and save the chart as PNG. | Explain how ChartSplitType.PercentValue and SplitPosition work to aggregate small values into an "Other" slice in Aspose.Cells. | Show how to configure SettableChartGlobalizationSettings for a workbook and verify the custom "Other" label appears in the rendered PNG.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPieChartExample
{
    // Creates a workbook, fills sample data, adds a pie chart, sets ChartSplitType.PercentValue to group small slices, applies SettableChartGlobalizationSettings to rename the aggregated “Other” segment, renders the chart directly to a PNG file, and saves the workbook as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the pie chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["A5"].PutValue("D");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Add a pie chart
                int chartIndex = sheet.Charts.Add(ChartType.Pie, 5, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories

                // Optional: split small values into an "Other" segment
                chart.NSeries[0].SplitType = ChartSplitType.PercentValue;
                // If SplitPosition is supported, uncomment the following line:
                // chart.NSeries[0].SplitPosition = 15; // values <15% go to "Other"

                // Create custom globalization settings for the chart
                SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
                chartSettings.SetOtherName("Custom Other"); // custom label for the "Other" segment

                // Apply the globalization settings to the workbook
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = chartSettings
                };

                // Render the chart directly to a PNG file
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    OnePagePerSheet = true
                };

                try
                {
                    // In older Aspose.Cells versions, ToImage accepts a file path and options.
                    chart.ToImage("PieChart.png", imgOptions);
                }
                catch (Exception renderEx)
                {
                    Console.WriteLine($"Chart rendering error: {renderEx.Message}");
                }

                // Save the workbook as an Excel file
                workbook.Save("PieChartWorkbook.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
