// Title: How to detect and add a secondary value axis to a column chart using Aspose.Cells in C#
// AI Prompts: Write C# code that uses Aspose.Cells to check if a chart already has a secondary value axis before setting PlotOnSecondAxis. | Show a C# example that conditionally creates a visible secondary axis for the second series in an Aspose.Cells column chart. | Demonstrate how to use chart.HasAxis with AxisType.Value and false to determine secondary axis presence in Aspose.Cells.
// Common Searches: asp.net aspose.cells check if chart has secondary axis before adding | c# aspose.cells column chart detect secondary value axis | how to use chart.HasAxis to verify secondary axis in Aspose.Cells | conditionally enable secondary axis for second series aspose.cells | aspose.cells make secondary value axis visible c# example
// Tags: Aspose.Cells chart.HasAxis secondary axis detection | Aspose.Cells PlotOnSecondAxis conditional usage | Aspose.Cells secondary value axis visibility | C# Aspose.Cells column chart secondary axis | Aspose.Cells add secondary axis only if missing

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds sample data, inserts a column chart, checks for an existing secondary value axis with chart.HasAxis, and if none is found, plots the second series on the secondary axis, makes the axis visible, and saves the workbook.
    public class CheckSecondaryAxisDemo
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

                // Add two series to the chart
                chart.NSeries.Add("B2:B4", true); // Series 1
                chart.NSeries.Add("C2:C4", true); // Series 2
                chart.NSeries.CategoryData = "A2:A4";

                // Check if a secondary value axis already exists
                bool hasSecondaryValueAxis = chart.HasAxis(AxisType.Value, false);
                Console.WriteLine("Secondary Value Axis exists before modification: " + hasSecondaryValueAxis);

                // If it does not exist, enable secondary axis for the second series
                if (!hasSecondaryValueAxis)
                {
                    // Plot the second series on the secondary value axis
                    chart.NSeries[1].PlotOnSecondAxis = true;

                    // Optionally customize the secondary axis
                    chart.SecondValueAxis.IsVisible = true;
                    chart.SecondValueAxis.Title.Text = "Secondary Axis";
                }

                // Verify again after modification
                bool hasSecondaryValueAxisAfter = chart.HasAxis(AxisType.Value, false);
                Console.WriteLine("Secondary Value Axis exists after modification: " + hasSecondaryValueAxisAfter);

                // Save the workbook
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CheckSecondaryAxisDemo.xlsx");
                workbook.Save(outputPath);
                Console.WriteLine("Workbook saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CheckSecondaryAxisDemo.Run();
        }
    }
}
