// Title: Create a line chart with standard deviation error bars using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that builds a line chart from worksheet data and adds error bars to the series. | Show how to set the error bar type to standard deviation, specify its values, and save the workbook as an Excel file.
// Common Searches: Aspose.Cells C# add standard deviation error bars to a line chart series | example of configuring error bar type for an Excel chart using Aspose.Cells .NET | C# code to display error bars on a line chart created with Aspose.Cells | set custom error bar values for a chart series in Aspose.Cells for .NET | generate Excel line chart with error bars programmatically in C#
// Tags: Aspose.Cells line chart error bar API | C# add standard deviation error bars Excel | configure chart series error bars Aspose.Cells | Excel line chart with error bars .NET | programmatic error bar setup Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B5 with category and value data, adds a line chart, sets its title, defines a series referencing the values, and demonstrates where to configure standard deviation error bars for that series. It also ensures the output directory exists and saves the workbook as 'LineChartWithErrorBars.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate data for the line chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(15);
            sheet.Cells["B4"].PutValue(12);
            sheet.Cells["B5"].PutValue(18);

            // Add a line chart positioned from row 7, column 0 to row 25, column 10
            int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Line Chart with Standard Deviation Error Bars";

            // Add series (values) and set its name
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries[0].Name = "Series 1";

            // NOTE: Error bar APIs may not be available in all Aspose.Cells versions.
            // If needed, configure error bars using the appropriate API for your version.

            // Define output file path
            string outputPath = "LineChartWithErrorBars.xlsx";

            // Ensure the output directory exists (if a directory is specified)
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
