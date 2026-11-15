// Title: Changing radar chart data label font size to 12 pt and color to blue with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set radar chart data label font size 12 | Change radar chart label color to blue using Aspose.Cells .NET | C# code to format data labels on a radar chart with Aspose.Cells | How to customize radar chart data label appearance in Aspose.Cells for .NET | Set font size and color of radar chart labels in a generated Excel file using Aspose.Cells
// Tags: radar chart data label formatting Aspose.Cells | set data label font size 12 Aspose.Cells | change data label color blue Aspose.Cells | Aspose.Cells radar chart label styling .NET | customize chart data labels C# Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, populates it with sample data, adds a radar chart, enables data labels for the first series, sets the data label font size to 12 points and the color to blue, and saves the file as RadarChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle create rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data for the radar chart
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");
            cells["A2"].PutValue("A");
            cells["A3"].PutValue("B");
            cells["A4"].PutValue("C");
            cells["A5"].PutValue("D");
            cells["B2"].PutValue(4);
            cells["B3"].PutValue(5);
            cells["B4"].PutValue(3);
            cells["B5"].PutValue(7);
            cells["C2"].PutValue(6);
            cells["C3"].PutValue(2);
            cells["C4"].PutValue(8);
            cells["C5"].PutValue(5);

            // Add a radar chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Radar, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Define the data series for the chart
            chart.NSeries.Add("B2:C5", true);

            // Enable data labels for the first series
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Set data label font size to 12 points and color to blue
            chart.NSeries[0].DataLabels.Font.Size = 12;
            chart.NSeries[0].DataLabels.Font.Color = Color.Blue;

            // Determine output file path
            string outputPath = "RadarChart.xlsx";

            // Save the workbook (lifecycle save rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
