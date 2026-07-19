// Title: Aspose.Cells C# – Apply Percentage Number Format to the Third Series Data Labels in a Column Chart
// Description: Creates a workbook, adds three data series, inserts a column chart, enables data labels for all series, and formats the third series' labels to show percentages with two decimal places using the "0.00%" NumberFormat and ShowPercentage property.
// Keywords: Aspose.Cells C# chart data label format | percentage data labels Aspose.Cells | format third series chart Aspose.Cells | C# set number format chart series | Aspose.Cells column chart percentage labels
// Common Searches: Aspose.Cells format data labels as percent for specific series | C# Aspose.Cells column chart third series percentage label | How to set number format for chart series data labels in Aspose.Cells | ShowPercentage property Aspose.Cells C# example
// Developer Intent: Apply a percentage number format to the data labels of the third series in a column chart using Aspose.Cells for .NET.
// Use Cases: Sales dashboard where the third series represents market‑share and must display percent labels. | Financial report showing profit‑margin as a ratio that needs two‑decimal‑place percentage labels. | KPI workbook where a ratio metric is visualized in a chart and requires percent formatting on its data labels.
// AI Prompts: Generate C# code to set a "0.00%" number format on the third series data labels of an Aspose.Cells column chart. | Explain how to enable ShowPercentage and apply a custom percentage format to a specific chart series in Aspose.Cells. | Provide a step‑by‑step example for formatting only one series' data labels as percentages in an Aspose.Cells workbook.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds three data series, inserts a column chart, enables data labels for all series, and formats the third series' labels to show percentages with two decimal places using the "0.00%" NumberFormat and ShowPercentage property.
    public class ApplyPercentageFormatToThirdSeries
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the three series to the chart
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.Add("D2:D4", true); // Series 3

            // Set category (X) data
            chart.NSeries.CategoryData = "A2:A4";

            // Enable data labels for all series
            foreach (Series s in chart.NSeries)
            {
                s.DataLabels.ShowValue = true;
            }

            // Apply percentage number format to data labels of the third series (index 2)
            Series thirdSeries = chart.NSeries[2];
            thirdSeries.DataLabels.NumberFormat = "0.00%";   // Two decimal places as percent
            thirdSeries.DataLabels.ShowPercentage = true;   // Ensure percentage is displayed

            // Determine output file path
            string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "ApplyPercentageFormatToThirdSeries.xlsx");

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}
