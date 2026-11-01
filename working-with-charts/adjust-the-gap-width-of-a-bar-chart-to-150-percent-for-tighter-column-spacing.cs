// Title: Adjust a clustered column chart’s gap width to 150 % for tighter spacing with Aspose.Cells for .NET (C#)
// AI Prompts: Create a new workbook, add sample data, insert a Column chart, and assign 150 to the chart.GapWidth property. | Open an existing workbook, locate a chart object, and modify its column spacing by setting chart.GapWidth = 150 in C#. | Save the workbook containing the modified chart as an .xlsx file and output the full file path.
// Common Searches: Aspose.Cells how to change column chart spacing to 150 percent in C# | C# code for reducing gap between bars in an Excel clustered column chart using Aspose | set chart.GapWidth to 150 Aspose.Cells .NET example | tighten column spacing in generated Excel chart programmatically
// Tags: Aspose.Cells chart GapWidth adjustment | C# set column chart spacing percentage | Excel clustered column chart gap width property | Aspose.Cells generate chart with custom spacing | programmatic column chart layout Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Shows how to create a workbook, add data, generate a clustered column chart, set its GapWidth to 150 % for tighter column spacing, and save the result as BarChartWithAdjustedGapWidth.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a clustered column chart (vertical bars)
            int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Adjust the gap width to 150 percent for tighter column spacing
            chart.GapWidth = 150;

            // Define output file path
            string outputPath = "BarChartWithAdjustedGapWidth.xlsx";

            // Ensure the output directory exists (if any)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
