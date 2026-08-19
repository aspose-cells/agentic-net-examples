// Title: Aspose.Cells for .NET: Hide Tick Marks on the Secondary Y‑Axis in a Mixed Column Chart
// Description: This C# example creates a workbook, adds category data and two series, builds a column chart, plots the second series on the secondary Y‑axis, and removes both major and minor tick marks from that axis by setting Axis.MajorTickMark and Axis.MinorTickMark to TickMarkType.None before saving the file.
// Keywords: Aspose.Cells hide secondary Y axis tick marks | C# Aspose.Cells chart tick marks | remove secondary axis ticks | TickMarkType.None | Aspose.Cells chart formatting | secondary value axis styling | mixed chart axis customization | Excel chart hide ticks programmatically | Aspose.Cells .NET chart API
// Common Searches: hide secondary Y axis tick marks Aspose.Cells | C# Aspose.Cells set secondary axis major tick mark to none | remove minor tick marks from secondary axis in Excel using Aspose | Aspose.Cells chart axis tick mark options | how to hide secondary axis ticks in .NET
// Developer Intent: Programmatically suppress major and minor tick marks on the secondary Y‑axis of an Excel chart.
// Use Cases: Generate a sales‑vs‑budget mixed column chart where the secondary axis shows revenue in thousands but its tick marks are hidden for a cleaner visual. | Create automated Excel reports that use a secondary scale without cluttering the axis with tick marks. | Apply consistent chart styling in bulk by disabling secondary axis tick marks across multiple workbooks.
// AI Prompts: Show C# code using Aspose.Cells to set the secondary value axis MajorTickMark and MinorTickMark to TickMarkType.None. | Give a step‑by‑step guide for adding a secondary series to a column chart and hiding its axis tick marks with Aspose.Cells for .NET. | Explain how to access chart.SecondValueAxis in Aspose.Cells and customize its tick mark visibility.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This C# example creates a workbook, adds category data and two series, builds a column chart, plots the second series on the secondary Y‑axis, and removes both major and minor tick marks from that axis by setting Axis.MajorTickMark and Axis.MinorTickMark to TickMarkType.None before saving the file.
    public class HideSecondaryYAxisTickMarks
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

                // Series 1 (primary axis)
                worksheet.Cells["B1"].PutValue("Primary");
                worksheet.Cells["B2"].PutValue(100);
                worksheet.Cells["B3"].PutValue(200);
                worksheet.Cells["B4"].PutValue(300);

                // Series 2 (secondary axis)
                worksheet.Cells["C1"].PutValue("Secondary");
                worksheet.Cells["C2"].PutValue(5000);
                worksheet.Cells["C3"].PutValue(3000);
                worksheet.Cells["C4"].PutValue(1000);

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Add the two series
                chart.NSeries.Add("B2:B4", true); // primary series
                chart.NSeries.Add("C2:C4", true); // secondary series
                chart.NSeries.CategoryData = "A2:A4";

                // Plot the second series on the secondary Y axis
                chart.NSeries[1].PlotOnSecondAxis = true;

                // Hide tick marks on the secondary Y axis
                Axis secondaryValueAxis = chart.SecondValueAxis;
                secondaryValueAxis.MajorTickMark = TickMarkType.None; // hide major tick marks
                secondaryValueAxis.MinorTickMark = TickMarkType.None; // hide minor tick marks

                // Save the workbook
                string outputPath = "HideSecondaryYAxisTickMarks.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
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
            HideSecondaryYAxisTickMarks.Run();
        }
    }
}
