// Title: Increase bar spacing in a column chart by setting the GapWidth property using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and sets its GapWidth to 200 to widen the space between bars. | Show how to adjust the GapWidth of a column chart's primary axis in Aspose.Cells and save the workbook. | Provide an example that modifies column series spacing via the Chart.GapWidth property in a .NET application.
// Common Searches: Aspose.Cells C# how to change column chart bar spacing | set GapWidth for column chart in Aspose.Cells .NET example | increase column series gap width primary axis Aspose.Cells | adjust spacing between columns in a chart using Aspose.Cells C# | Aspose.Cells chart GapWidth property usage tutorial
// Tags: Aspose.Cells column chart gap width | C# set chart bar spacing Aspose.Cells | Chart.GapWidth property .NET | increase column series spacing Aspose.Cells | adjust column chart spacing programmatically

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, adds sample data, inserts a column chart, sets its data range, changes the chart's GapWidth to 200 to increase spacing between bars, and saves the file as AdjustedGapWidth.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the column chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart to the worksheet (from row 5, column 0 to row 20, column 10)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (true = columns as categories)
            chart.NSeries.Add("B2:C4", true);

            // Adjust the gap width of the column series on the primary axis
            // GapWidth is expressed as a percentage (default is 150). Higher values increase spacing.
            chart.GapWidth = 200; // Example: increase spacing between bars

            // Determine output file path and ensure the directory exists
            string outputPath = "AdjustedGapWidth.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
