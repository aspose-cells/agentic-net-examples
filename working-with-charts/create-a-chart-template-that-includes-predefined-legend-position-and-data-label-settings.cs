// Title: Generate an Excel column chart template with a bottom legend and data labels (values and categories) using Aspose.Cells for .NET
// AI Prompts: Create a new Workbook, add sample data, insert a column chart, set the legend position to Bottom, enable automatic sizing, and apply a bold 11‑pt font to the legend. | For each series in the chart, enable DataLabels and configure them to display both the value and the category name. | Calculate the chart layout and save the workbook as an .xlsx file named ChartWithLegendAndDataLabels.xlsx.
// Common Searches: Aspose.Cells how to place chart legend at the bottom and make it bold | C# set data labels to display value and category name for multiple series in Aspose.Cells chart | Create reusable chart template with predefined legend settings using Aspose.Cells .NET | Save an Excel workbook with a column chart that has automatic legend sizing in Aspose.Cells | Example code for configuring legend font size and data labels in Aspose.Cells chart
// Tags: Aspose.Cells column chart legend bottom placement | Aspose.Cells chart series data label configuration | Aspose.Cells automatic legend sizing and bold font | Aspose.Cells save workbook as xlsx with chart | Aspose.Cells reusable chart template .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ChartTemplateExample
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, configures the legend to appear at the bottom with automatic sizing and a bold 11‑pt font, enables data labels that show both values and category names for each series, recalculates the chart layout, and saves the result as ChartWithLegendAndDataLabels.xlsx.
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
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 1, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart (including categories and series)
                chart.NSeries.Add("B2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Predefined Legend Settings
                chart.Legend.Position = LegendPositionType.Bottom;
                // Overlay setting not available in this version; omitted.
                chart.Legend.IsAutomaticSize = true;
                chart.Legend.Font.Size = 11;
                chart.Legend.Font.IsBold = true;

                // Predefined Data Label Settings
                // Enable data labels for the first series
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[0].DataLabels.ShowCategoryName = true;
                // Position setting not available in this version; omitted.

                // Enable data labels for the second series
                chart.NSeries[1].DataLabels.ShowValue = true;
                chart.NSeries[1].DataLabels.ShowCategoryName = true;
                // Position setting not available in this version; omitted.

                // Calculate the chart to apply layout changes
                chart.Calculate();

                // Define output file path
                string outputPath = "ChartWithLegendAndDataLabels.xlsx";

                // Save the workbook with the configured chart
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
