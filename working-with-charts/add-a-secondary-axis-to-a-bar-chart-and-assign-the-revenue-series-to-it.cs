// Title: C# – Build a 2‑D Bar Chart with a Secondary Y‑Axis (Revenue) using Aspose.Cells
// Description: The sample creates a workbook, populates month, revenue and cost data, adds a 2‑D bar chart, places the cost series on the primary axis and the revenue series on a secondary Y‑axis, customizes axis titles and range, and saves the result as BarChartWithSecondaryAxis.xlsx.
// Keywords: Aspose.Cells C# secondary axis | dual axis bar chart Aspose.Cells | plot revenue on secondary Y axis | Aspose.Cells chart customization | add secondary value axis .NET | BarChartWithSecondaryAxis example | Aspose.Cells set axis title | configure axis min max Aspose.Cells | Excel financial chart C# | move series to secondary axis Aspose.Cells
// Common Searches: Aspose.Cells add secondary axis to bar chart | C# plot series on secondary Y axis Aspose.Cells | set secondary axis title in Aspose.Cells chart | dual axis chart example Aspose.Cells .NET | bar chart with cost and revenue separate scales
// Developer Intent: Add a secondary Y‑axis to a bar chart and bind the revenue series to it.
// Use Cases: Show cost and revenue together while each uses its own scale. | Automate quarterly financial dashboards that require distinct axes for different metrics. | Create marketing‑spend vs. sales visuals where values differ in magnitude.
// AI Prompts: Generate C# code with Aspose.Cells that creates a bar chart, adds a secondary Y‑axis, sets axis titles, and defines custom min/max values. | Explain how to move a specific series to the secondary axis in an Aspose.Cells chart and adjust its formatting. | Show the steps to configure the secondary value axis (title, minimum, maximum, major unit) in Aspose.Cells. | Provide a concise guide to export a chart with dual axes to an Excel file using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The sample creates a workbook, populates month, revenue and cost data, adds a 2‑D bar chart, places the cost series on the primary axis and the revenue series on a secondary Y‑axis, customizes axis titles and range, and saves the result as BarChartWithSecondaryAxis.xlsx.
    public class BarChartWithSecondaryAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                // Column A: Categories
                worksheet.Cells["A1"].PutValue("Month");
                worksheet.Cells["A2"].PutValue("Jan");
                worksheet.Cells["A3"].PutValue("Feb");
                worksheet.Cells["A4"].PutValue("Mar");

                // Column B: Revenue (will be plotted on secondary axis)
                worksheet.Cells["B1"].PutValue("Revenue");
                worksheet.Cells["B2"].PutValue(120000);
                worksheet.Cells["B3"].PutValue(150000);
                worksheet.Cells["B4"].PutValue(130000);

                // Column C: Cost (primary axis)
                worksheet.Cells["C1"].PutValue("Cost");
                worksheet.Cells["C2"].PutValue(80000);
                worksheet.Cells["C3"].PutValue(90000);
                worksheet.Cells["C4"].PutValue(85000);

                // Add a 2‑D bar chart
                int chartIndex = worksheet.Charts.Add(ChartType.Bar, 6, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the Cost series (primary axis)
                chart.NSeries.Add("C2:C4", true);
                // Add the Revenue series (will be moved to secondary axis)
                chart.NSeries.Add("B2:B4", true);

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the Revenue series (second series, index 1) on the secondary Y axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Optional: customize the secondary value axis (e.g., title and range)
                Axis secondaryAxis = chart.SecondValueAxis;
                secondaryAxis.Title.Text = "Revenue (USD)";
                secondaryAxis.MinValue = 0;
                secondaryAxis.MaxValue = 200000;
                secondaryAxis.MajorUnit = 50000;

                // Optional: customize the primary value axis for Cost
                chart.ValueAxis.Title.Text = "Cost (USD)";

                // Save the workbook
                string outputPath = "BarChartWithSecondaryAxis.xlsx";
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
            BarChartWithSecondaryAxis.Run();
        }
    }
}
