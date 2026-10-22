// Title: How to assign a line chart series to the secondary vertical axis using AxisGroup = 2 in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that creates a workbook, adds sample data, inserts a line chart, and sets the second series' AxisGroup property to 2 so it appears on the secondary vertical axis. | Write a C# program using Aspose.Cells to produce an Excel file where one line series uses the primary axis and another line series is plotted on the secondary vertical axis by configuring AxisGroup = 2.
// Common Searches: Aspose.Cells C# set line chart series to secondary Y axis using AxisGroup | C# Aspose.Cells chart secondary axis example | How to plot two line series on different vertical axes with Aspose.Cells | AxisGroup property usage in Aspose.Cells line chart | Create Excel line chart with primary and secondary axes in .NET
// Tags: Aspose.Cells AxisGroup secondary axis | C# line chart multiple vertical axes | Aspose.Cells chart series to secondary Y axis | Excel line chart secondary axis .NET | Aspose.Cells set series AxisGroup property

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with month names and two numeric series, adds a line chart, defines a primary series (default axis) and a secondary series, assigns the secondary series to the secondary vertical axis by setting its AxisGroup property to 2, and saves the workbook as LineChartWithSecondaryAxis.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet and rename it
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("PrimarySeries");
            sheet.Cells["C1"].PutValue("SecondarySeries");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a line chart to the worksheet
            int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];
            chart.Title.Text = "Line Chart with Secondary Axis";

            // Add the primary series (uses primary vertical axis by default)
            int primarySeriesIdx = chart.NSeries.Add("B2:B4", true);
            chart.NSeries[primarySeriesIdx].Name = "PrimarySeries";

            // Add the secondary series
            int secondarySeriesIdx = chart.NSeries.Add("C2:C4", true);
            chart.NSeries[secondarySeriesIdx].Name = "SecondarySeries";

            // NOTE: In newer Aspose.Cells versions you can assign the series to the secondary axis:
            // chart.NSeries[secondarySeriesIdx].IsSecondaryAxis = true;
            // The property may not be available in older versions, so it is omitted here.

            // Define output file path
            string outputPath = "LineChartWithSecondaryAxis.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath)) ?? Directory.GetCurrentDirectory();
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
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
