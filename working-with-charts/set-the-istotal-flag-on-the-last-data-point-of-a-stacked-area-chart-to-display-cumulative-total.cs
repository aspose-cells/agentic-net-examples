// Title: Aspose.Cells for .NET – Set IsTotal on Final Point of a Stacked Area Chart
// Description: Creates a workbook, adds sample data, builds a stacked area chart, and applies LayoutProperties.Subtotals to label the last point as the overall total before saving as XLSX.
// Keywords: Aspose.Cells | .NET | stacked area chart | IsTotal flag | LayoutProperties.Subtotals | cumulative total | chart data point | C# example | Excel automation | chart subtotal
// Common Searches: Aspose.Cells mark last chart point as total | C# set IsTotal flag stacked area | LayoutProperties.Subtotals usage | how to show cumulative total in Excel chart with Aspose | stacked area chart subtotal index .NET
// Developer Intent: Flag the final data point of a stacked area chart as a cumulative total.
// Use Cases: Financial statements where the month‑end total is highlighted in a stacked area visualization. | Quarterly KPI dashboards that need the overall total displayed for the last period. | Automated reporting scripts that generate Excel workbooks with cumulative totals shown in area charts.
// AI Prompts: Provide C# code that sets the IsTotal flag on the last point of a stacked area chart using Aspose.Cells. | Explain how LayoutProperties.Subtotals can be used to display a cumulative total in an Excel stacked area chart with Aspose.Cells for .NET. | Show an example of calculating the correct point index and applying the Subtotals property to a chart series in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds sample data, builds a stacked area chart, and applies LayoutProperties.Subtotals to label the last point as the overall total before saving as XLSX.
    public class StackedAreaChartIsTotalDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Populate sample data for a stacked area chart
                // -------------------------------------------------
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["D1"].PutValue("Series3");

                // Data rows
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);
                sheet.Cells["C5"].PutValue(45);

                sheet.Cells["D2"].PutValue(5);
                sheet.Cells["D3"].PutValue(10);
                sheet.Cells["D4"].PutValue(15);
                sheet.Cells["D5"].PutValue(20);

                // -------------------------------------------------
                // Add a stacked area chart
                // -------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.AreaStacked, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add each series (vertical = true because data are in columns)
                chart.NSeries.Add("B2:B5", true); // Series1
                chart.NSeries.Add("C2:C5", true); // Series2
                chart.NSeries.Add("D2:D5", true); // Series3

                // Set category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A5";

                // -------------------------------------------------
                // Mark the last data point as a total (cumulative)
                // -------------------------------------------------
                // The last point index is zero‑based; for 4 data rows it is 3
                int lastPointIndex = 3;
                // Apply the subtotal index to the first series (any series works for stacked charts)
                chart.NSeries[0].LayoutProperties.Subtotals = new int[] { lastPointIndex };

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "StackedAreaChartIsTotalDemo.xlsx";

                // Ensure the directory exists (useful if a relative path with folders is provided)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the stacked area chart:");
                Console.WriteLine(ex.Message);
            }
        }
    }

    // Entry point for the console application
    public class Program
    {
        public static void Main(string[] args)
        {
            StackedAreaChartIsTotalDemo.Run();
        }
    }
}
