// Title: Format third‑series data labels as percentages in an Aspose.Cells column chart (C#)
// Description: Creates a workbook with three series, adds a column chart, enables data labels for the third series, shows value and percentage, and applies the "0.00%" number format to the labels before saving.
// Keywords: Aspose.Cells C# chart data label format | percentage number format series 3 | set data label format column chart Aspose | .NET Excel chart custom number format | format third series labels as percent
// Common Searches: Aspose.Cells set percentage format for specific series | C# chart data labels custom number format Aspose | how to show percent on third series in Excel chart using Aspose | apply 0.00% format to chart labels .NET
// Developer Intent: Apply a custom percentage number format to the data labels of the third series in an Aspose.Cells column chart.
// Use Cases: Financial reports where the third series represents profit margin and must display percentages with two decimals. | Sales dashboards that show market‑share data in the third series and need clear percentage labels. | Automated Excel exports that require the third series of a chart to be labeled as "0.00%" for presentation consistency.
// AI Prompts: Generate C# code with Aspose.Cells that formats the third series data labels as percentages using "0.00%". | Explain how to enable both value and percentage display for a specific chart series and apply a custom number format. | Adapt the example to a line chart while keeping the percentage data label formatting for the third series.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with three series, adds a column chart, enables data labels for the third series, shows value and percentage, and applies the "0.00%" number format to the labels before saving.
    public class ThirdSeriesDataLabelPercentageFormat
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
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");

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

                // Series 3 values (the one we will format)
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

                // Enable data labels for the third series
                Series thirdSeries = chart.NSeries[2]; // zero‑based index
                thirdSeries.DataLabels.ShowValue = true;          // show the raw value
                thirdSeries.DataLabels.ShowPercentage = true;    // show percentage
                // Apply a percentage number format to the data labels
                thirdSeries.DataLabels.NumberFormat = "0.00%";

                // Save the workbook
                string outputPath = "ThirdSeriesDataLabelPercentageFormat.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            ThirdSeriesDataLabelPercentageFormat.Run();
        }
    }
}
