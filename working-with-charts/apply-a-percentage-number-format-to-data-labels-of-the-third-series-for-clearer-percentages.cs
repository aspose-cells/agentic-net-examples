// Title: How to format the third series data labels as percentages in an Aspose.Cells column chart using C#
// AI Prompts: Write C# code that adds data labels to the third series of an Aspose.Cells column chart and sets the number format to "0.00%". | Show how to enable percentage display for a specific series in an Aspose.Cells chart and customize the label format. | Provide a C# snippet that configures a column chart's third series to show both the raw value and its percentage with a custom number format.
// Common Searches: Aspose.Cells C# set percentage number format for data labels of a specific chart series | How to display 0.00% format on third series data labels in an Aspose.Cells column chart | C# Aspose.Cells chart show value and percentage in data labels for one series | Apply custom number format to series data labels in Aspose.Cells workbook
// Tags: Aspose.Cells chart series data label percentage format | C# set number format for chart data labels | column chart third series data labels Aspose.Cells | custom number format 0.00% Aspose.Cells | enable data labels for specific series Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds three data series, builds a column chart, enables data labels for the third series, applies a "0.00%" number format to show percentages, and saves the file as an Excel workbook.
    public class ApplyPercentageFormatToThirdSeriesDataLabels
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for three series
                // Categories
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                // Series 1 values
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Series 2 values
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Series 3 values (to be shown as percentages)
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(0.1);   // 10%
                sheet.Cells["D3"].PutValue(0.2);   // 20%
                sheet.Cells["D4"].PutValue(0.3);   // 30%

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Add the three series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.Add("D2:D4", true); // Series 3

                // Set category (X) data
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the third series and apply percentage format
                Series thirdSeries = chart.NSeries[2]; // zero‑based index
                thirdSeries.DataLabels.ShowValue = true;          // show the value
                thirdSeries.DataLabels.ShowPercentage = true;    // optional: display percentage
                thirdSeries.DataLabels.NumberFormat = "0.00%";   // apply percentage number format

                // Save the workbook
                string outputPath = "ThirdSeriesPercentageDataLabels.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            Run();
        }
    }
}
