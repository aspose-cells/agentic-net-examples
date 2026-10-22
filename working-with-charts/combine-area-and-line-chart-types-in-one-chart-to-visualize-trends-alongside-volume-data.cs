// Title: Create a mixed Area‑Line chart with month categories in an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates an Excel workbook, inserts an Area series for volume data and a Line series for trend data on the same chart, and sets month names as the category axis with Aspose.Cells. | Show how to change the chart type of individual series after they have been added to an Aspose.Cells chart in C#. | Demonstrate exporting the workbook that contains both area and line series to an .xlsx file using Aspose.Cells.
// Common Searches: asp.net aspose.cells create combined area line chart with month category axis | c# set different chart types for each series in Aspose.Cells chart | how to add a line series to an existing area chart using Aspose.Cells | example of area and line series together for monthly data with Aspose.Cells .NET
// Tags: insert area series Aspose.Cells C# | insert line series Aspose.Cells C# | modify series chart type Aspose.Cells | configure category axis month labels Aspose.Cells | save workbook to xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The sample creates a new workbook, fills columns with month, volume, and trend values, adds an Area chart, inserts an Area series for volume and a Line series for trend, assigns the month range as the category axis, and saves the file as CombinedAreaLineChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data: Month, Volume (area), Trend (line)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Volume");
            sheet.Cells["C1"].PutValue("Trend");

            string[] months = { "Jan", "Feb", "Mar", "Apr", "May", "Jun" };
            double[] volume = { 120, 150, 130, 170, 160, 180 };
            double[] trend = { 100, 130, 110, 150, 140, 160 };

            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A: Month
                sheet.Cells[i + 1, 1].PutValue(volume[i]);  // Column B: Volume
                sheet.Cells[i + 1, 2].PutValue(trend[i]);   // Column C: Trend
            }

            // Add a chart to the worksheet (initially Area type)
            int chartIndex = sheet.Charts.Add(ChartType.Area, 8, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set chart title
            chart.Title.Text = "Volume (Area) and Trend (Line)";

            // Add Area series for Volume (B2:B7)
            chart.NSeries.Add("B2:B7", true);
            chart.NSeries[0].Name = "Volume";
            chart.NSeries[0].Type = ChartType.Area; // Ensure this series is Area

            // Add Line series for Trend (C2:C7)
            chart.NSeries.Add("C2:C7", true);
            chart.NSeries[1].Name = "Trend";
            chart.NSeries[1].Type = ChartType.Line; // Set this series to Line

            // Set category axis labels (Month column)
            chart.NSeries.CategoryData = "A2:A7";

            // Save the workbook with the combined chart
            workbook.Save("CombinedAreaLineChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
