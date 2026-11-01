// Title: Create a column chart with Aspose.Cells for .NET and automatically resize the right‑hand legend based on the longest series name
// AI Prompts: Write C# code that uses Aspose.Cells to add a column chart, place the legend on the right side, and set the legend width according to the length of the longest series label. | Demonstrate how to compute the required legend width in points and apply it to an Aspose.Cells chart after populating series data. | Show how to adjust both legend width and height programmatically for an Excel chart generated with Aspose.Cells.
// Common Searches: Aspose.Cells C# set legend width based on series name length | dynamic legend sizing for column chart using Aspose.Cells | how to programmatically adjust Excel chart legend dimensions in .NET | calculate legend size from series labels with Aspose.Cells | auto‑resize right‑hand legend in Aspose.Cells chart
// Tags: Aspose.Cells chart legend dynamic width | column chart legend positioning Aspose.Cells | programmatic legend size adjustment C# | calculate legend dimensions from series names | Excel chart auto resize legend Aspose.Cells

using System;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with sample data, adds a column chart, positions the legend on the right with an initial size, computes the longest series name, adjusts the legend width accordingly, and saves the workbook as an Excel file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set data series and categories
            chart.NSeries.Add("B2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Position legend on the right with an initial size
            chart.Legend.Position = LegendPositionType.Right;
            chart.Legend.Width = 100;   // initial width (points)
            chart.Legend.Height = 50;   // initial height (points)

            // Dynamically adjust legend width based on the longest series name
            var seriesNames = chart.NSeries.Select(s => s.Name).ToList();
            if (seriesNames.Count > 0)
            {
                int maxLength = seriesNames.Max(name => name?.Length ?? 0);
                // Approximate width: 7 points per character + padding
                chart.Legend.Width = maxLength * 7 + 20;
            }

            // Save the workbook
            string outputPath = "ChartWithDynamicLegend.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
