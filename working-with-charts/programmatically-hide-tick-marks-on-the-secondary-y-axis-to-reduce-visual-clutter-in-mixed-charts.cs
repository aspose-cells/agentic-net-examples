// Title: How to hide tick marks on the secondary Y‑axis of a mixed column‑line chart with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that builds a mixed column and line chart, assigns the line series to the secondary Y‑axis, and disables both major and minor tick marks on that axis. | Demonstrate how to retrieve the Chart.SecondValueAxis in Aspose.Cells and set its MajorTickMark and MinorTickMark properties to TickMarkType.None before saving the workbook.
// Common Searches: Aspose.Cells C# hide secondary axis tick marks in mixed column line chart | remove major and minor tick marks from secondary Y axis using Aspose.Cells .NET | how to disable tick marks on secondary value axis in Aspose.Cells chart | mixed chart secondary Y axis formatting Aspose.Cells C# example
// Tags: secondary-value-axis tick-mark removal Aspose.Cells | mixed column line chart axis customization .NET | set MajorTickMark None Aspose.Cells | hide minor tick marks Aspose.Cells chart | Excel chart visual clutter reduction Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a workbook, fills it with sample data, adds a mixed column‑line chart, plots the second series on the secondary Y‑axis, and hides both major and minor tick marks on that axis before saving the file as an XLSX workbook.
    class HideSecondaryYAxisTickMarks
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

                worksheet.Cells["B1"].PutValue("Primary");
                worksheet.Cells["B2"].PutValue(10);
                worksheet.Cells["B3"].PutValue(20);
                worksheet.Cells["B4"].PutValue(30);

                worksheet.Cells["C1"].PutValue("Secondary");
                worksheet.Cells["C2"].PutValue(1000);
                worksheet.Cells["C3"].PutValue(2000);
                worksheet.Cells["C4"].PutValue(3000);

                // Add a mixed chart (column for primary series, line for secondary series)
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add primary series (column)
                chart.NSeries.Add("B2:B4", true);
                // Add secondary series (line) and plot it on the secondary Y axis
                chart.NSeries.Add("C2:C4", true);
                chart.NSeries.CategoryData = "A2:A4";
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Hide tick marks on the secondary Y axis
                Axis secondaryValueAxis = chart.SecondValueAxis;
                secondaryValueAxis.MajorTickMark = TickMarkType.None; // Hide major tick marks
                secondaryValueAxis.MinorTickMark = TickMarkType.None; // Hide minor tick marks

                // Save the workbook
                string outputPath = "MixedChart_SecondaryYAxis_NoTickMarks.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }

        // Entry point for the application
        static void Main(string[] args)
        {
            Run();
        }
    }
}
