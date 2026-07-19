// Title: C# – Add a Secondary Y‑Axis to a Column Chart Using Aspose.Cells
// Description: Demonstrates how to create a workbook, insert category data and two series, generate a column chart, plot the second series on a secondary Y‑axis, customize the axis title and scale, and save the file as an Excel workbook.
// Keywords: Aspose.Cells secondary axis C# | column chart secondary Y axis | plot series on secondary axis Aspose.Cells | customize secondary value axis | Aspose.Cells chart example | C# Excel chart secondary axis
// Common Searches: Aspose.Cells add secondary Y axis to column chart | C# plot series on secondary axis in Excel | how to set secondary axis title Aspose.Cells | Aspose.Cells column chart with two Y axes | secondary value axis range Aspose.Cells C#
// Developer Intent: Create a column chart, assign one series to a secondary Y‑axis, and configure that axis programmatically.
// Use Cases: Display sales volume and revenue together, using revenue on a larger‑scale secondary axis. | Compare temperature (°C) and precipitation (mm) in a single chart with separate axes. | Show production count and defect rate where defect rate requires a different scale.
// AI Prompts: Generate C# code that adds a secondary Y‑axis to an Aspose.Cells column chart and assigns a specific series to it. | Explain how to set the title, minimum, maximum, and major unit of the secondary value axis in Aspose.Cells. | Show how to retrieve the secondary axis object after creating a chart and modify its properties programmatically.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, insert category data and two series, generate a column chart, plot the second series on a secondary Y‑axis, customize the axis title and scale, and save the file as an Excel workbook.
    public class SecondaryYAxisDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Series 1");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                worksheet.Cells["C1"].PutValue("Series 2");
                worksheet.Cells["C2"].PutValue(5000);
                worksheet.Cells["C3"].PutValue(3000);
                worksheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Add two data series
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary Y‑axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Customize the secondary Y‑axis (optional)
                Axis secondaryAxis = chart.SecondValueAxis;
                secondaryAxis.Title.Text = "Secondary Axis";
                secondaryAxis.MinValue = 0;
                secondaryAxis.MaxValue = 6000;
                secondaryAxis.MajorUnit = 1000;

                // Save the workbook
                workbook.Save("SecondaryYAxisDemo.xlsx");
                Console.WriteLine("Workbook saved as SecondaryYAxisDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            SecondaryYAxisDemo.Run();
        }
    }
}
