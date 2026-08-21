// Title: C# – Create a Pie Chart with a Custom “Other” Slice and Export to PNG using Aspose.Cells
// Description: This example shows how to build a workbook, fill it with category data, add a pie chart, group slices that represent less than 15 % of the total into an “Other” segment, rename that segment with SettableChartGlobalizationSettings, and render the chart as a high‑resolution PNG file. The workbook is also saved for reference.
// Keywords: Aspose.Cells pie chart C# | custom Other slice name | ChartSplitType PercentValue | ChartGlobalizationSettings | export chart to PNG | render Aspose.Cells chart | .NET spreadsheet chart image | localize Other segment Aspose
// Common Searches: rename Other slice in Aspose.Cells pie chart | set custom label for Other segment Aspose.Cells | group small values into Other slice C# | render Aspose.Cells pie chart as PNG | ChartGlobalizationSettings example .NET
// Developer Intent: Create a pie chart, combine minor categories into a custom‑named “Other” slice, and save the chart as a PNG image.
// Use Cases: Generate a sales‑distribution pie chart where categories below 15 % are merged into a “Miscellaneous Items” slice and embed the PNG in a quarterly report. | Produce a dashboard thumbnail that displays a localized “Other” segment, rendering the chart to PNG for web or mobile consumption. | Automate workbook creation that includes a pie chart with a custom “Other” label and export the chart image for email summaries or documentation.
// AI Prompts: Write C# code with Aspose.Cells to create a pie chart, set ChartSplitType.PercentValue, rename the Other slice to "Miscellaneous Items", and save the chart as a PNG file. | Explain how SettableChartGlobalizationSettings changes the label of the Other segment in Aspose.Cells pie charts and demonstrate its usage in a workbook. | Show how to configure ImageOrPrintOptions for 96 DPI resolution when rendering an Aspose.Cells chart to PNG.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Rendering;

namespace AsposeCellsPieChartExample
{
    // This example shows how to build a workbook, fill it with category data, add a pie chart, group slices that represent less than 15 % of the total into an “Other” segment, rename that segment with SettableChartGlobalizationSettings, and render the chart as a high‑resolution PNG file. The workbook is also saved for reference.
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

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Group small values into an "Other" segment (less than 15%)
                chart.NSeries[0].SplitType = ChartSplitType.PercentValue;

                // Customize the label for the "Other" segment
                SettableChartGlobalizationSettings chartSettings = new SettableChartGlobalizationSettings();
                chartSettings.SetOtherName("Miscellaneous Items");
                workbook.Settings.GlobalizationSettings = new GlobalizationSettings
                {
                    ChartSettings = chartSettings
                };

                // Prepare image rendering options (resolution)
                ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
                {
                    HorizontalResolution = 96,
                    VerticalResolution = 96
                };

                // Render the chart to a PNG file
                string chartImagePath = "PieChart.png";
                chart.ToImage(chartImagePath, imgOptions);

                // Save the workbook for reference
                string workbookPath = "PieChartWorkbook.xlsx";
                workbook.Save(workbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
