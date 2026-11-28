// Title: How to set a 2‑point border thickness on data labels of a stacked bar chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set data label border thickness for stacked bar chart | C# example to change data label border weight in Excel bar chart with Aspose.Cells | How to apply a 2 point border to chart data labels in Aspose.Cells .NET | Increase data label border size in a stacked bar chart using Aspose.Cells for .NET | Set data label border weight programmatically in Excel using Aspose.Cells C#
// Tags: chart data label border thickness Aspose.Cells | stacked bar chart label styling .NET | Excel chart border formatting programmatically C# | Aspose.Cells chart series label customization | data label outline weight Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;   // For line formatting

namespace SetDataLabelBorderThickness
{
    // The example creates a workbook, adds sample data, inserts a stacked bar chart, enables data labels for each series, sets the data label border weight to 2 points, and saves the file as an XLSX workbook using Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for a stacked bar chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series 1");
                sheet.Cells["C1"].PutValue("Series 2");

                sheet.Cells["A2"].PutValue("Item A");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["C2"].PutValue(20);

                sheet.Cells["A3"].PutValue("Item B");
                sheet.Cells["B3"].PutValue(40);
                sheet.Cells["C3"].PutValue(25);

                // Add a stacked bar chart
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:C3", true);          // Series values
                chart.NSeries.CategoryData = "A2:A3";      // Categories

                // Configure data labels for each series
                foreach (Series series in chart.NSeries)
                {
                    // Make data labels visible
                    series.DataLabels.ShowValue = true;

                    // Set the border thickness to 2 points
                    series.DataLabels.Border.WeightPt = 2.0;

                    // Optional: set border color
                    // series.DataLabels.Border.Color = System.Drawing.Color.Black;
                }

                // Save the workbook
                string outputPath = "StackedBarChart_WithDataLabelBorder.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
