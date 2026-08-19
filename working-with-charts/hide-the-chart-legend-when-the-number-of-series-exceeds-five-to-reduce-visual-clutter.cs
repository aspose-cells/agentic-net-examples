// Title: Hide Chart Legend in Aspose.Cells C# When Series Count Exceeds Five
// Description: This example creates a workbook, populates six data series, adds a column chart, and automatically hides the legend if the chart contains more than five series by checking the NSeries count and setting ShowLegend to false before saving the file.
// Keywords: Aspose.Cells chart legend | C# hide legend Aspose.Cells | conditional legend visibility | chart series count Aspose | column chart Aspose.Cells | ShowLegend property | .NET Excel chart example | reduce chart clutter
// Common Searches: Aspose.Cells hide legend when many series | C# chart legend conditional hide | How to hide Excel chart legend with Aspose | Hide legend if series count > 5 Aspose.Cells | Aspose.Cells column chart without legend
// Developer Intent: Automatically suppress the chart legend in an Aspose.Cells-generated Excel file when the number of data series exceeds a defined threshold.
// Use Cases: Sales dashboards with more than five product lines where the legend would overcrowd the view. | Financial reports that plot numerous categories and need a clean layout. | Automated Excel exports that adapt legend visibility based on dynamic data series counts.
// AI Prompts: Generate C# code using Aspose.Cells to hide a chart legend when the series count is greater than five. | Show how to check NSeries.Count and set ShowLegend = false for a column chart in Aspose.Cells. | Provide an Aspose.Cells example that conditionally removes the legend to improve readability of multi‑series charts.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example creates a workbook, populates six data series, adds a column chart, and automatically hides the legend if the chart contains more than five series by checking the NSeries count and setting ShowLegend to false before saving the file.
    public class HideLegendWhenManySeries
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook (lifecycle rule: create)
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for multiple series (6 series to trigger legend hide)
                // Category labels
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Q1");
                sheet.Cells["A3"].PutValue("Q2");
                sheet.Cells["A4"].PutValue("Q3");
                sheet.Cells["A5"].PutValue("Q4");

                // Series 1
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // Series 2
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                // Series 3
                sheet.Cells["D1"].PutValue("Series3");
                sheet.Cells["D2"].PutValue(12);
                sheet.Cells["D3"].PutValue(22);
                sheet.Cells["D4"].PutValue(32);
                sheet.Cells["D5"].PutValue(42);

                // Series 4
                sheet.Cells["E1"].PutValue("Series4");
                sheet.Cells["E2"].PutValue(18);
                sheet.Cells["E3"].PutValue(28);
                sheet.Cells["E4"].PutValue(38);
                sheet.Cells["E5"].PutValue(48);

                // Series 5
                sheet.Cells["F1"].PutValue("Series5");
                sheet.Cells["F2"].PutValue(14);
                sheet.Cells["F3"].PutValue(24);
                sheet.Cells["F4"].PutValue(34);
                sheet.Cells["F5"].PutValue(44);

                // Series 6 (extra series to exceed the threshold)
                sheet.Cells["G1"].PutValue("Series6");
                sheet.Cells["G2"].PutValue(16);
                sheet.Cells["G3"].PutValue(26);
                sheet.Cells["G4"].PutValue(36);
                sheet.Cells["G5"].PutValue(46);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add each series to the chart
                chart.NSeries.Add("B2:B5", true); // Series1
                chart.NSeries.Add("C2:C5", true); // Series2
                chart.NSeries.Add("D2:D5", true); // Series3
                chart.NSeries.Add("E2:E5", true); // Series4
                chart.NSeries.Add("F2:F5", true); // Series5
                chart.NSeries.Add("G2:G5", true); // Series6

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // Hide legend if the chart has more than five series
                if (chart.NSeries.Count > 5)
                {
                    chart.ShowLegend = false; // Hide legend to reduce visual clutter
                }

                // Determine output file path
                string outputPath = "ChartWithConditionalLegend.xlsx";

                // Save the workbook (lifecycle rule: save)
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            HideLegendWhenManySeries.Run();
        }
    }
}
