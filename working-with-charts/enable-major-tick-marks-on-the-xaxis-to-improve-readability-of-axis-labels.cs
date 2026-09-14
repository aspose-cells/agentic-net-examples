// Title: Enable major tick marks on the X‑axis of a column chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart using Aspose.Cells and sets the category axis major tick marks to Outside. | Show how to apply TickMarkType.Outside to the X‑axis of a chart in Aspose.Cells and then save the workbook. | Demonstrate configuring the CategoryAxis.MajorTickMark property for a column chart in a .NET application.
// Common Searches: Aspose.Cells C# set axis tick marks on chart X axis | configure category axis tick mark type to outside Aspose.Cells | C# example formatting chart axis tick marks with Aspose.Cells | enable axis tick marks on Excel chart using Aspose.Cells .NET | Aspose.Cells TickMarkType usage for chart axes
// Tags: Aspose.Cells chart category axis tick mark | C# column chart axis formatting Aspose.Cells | TickMarkType Outside usage Aspose.Cells | Excel chart axis customization Aspose.Cells | set X axis tick marks Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // // This example creates a new workbook, adds sample data, inserts a column chart, sets the X‑axis (category axis) major tick marks to Outside, and saves the file as EnableMajorTickMarksOnXAxis_out.xlsx.
    public class EnableMajorTickMarksOnXAxis
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable major tick marks on the X‑axis (category axis)
                chart.CategoryAxis.MajorTickMark = TickMarkType.Outside;

                // Save the workbook
                workbook.Save("EnableMajorTickMarksOnXAxis_out.xlsx");
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
            EnableMajorTickMarksOnXAxis.Run();
        }
    }
}
