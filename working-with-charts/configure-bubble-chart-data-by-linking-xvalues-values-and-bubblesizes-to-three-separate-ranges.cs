// Title: Binding XValues, Y Values, and BubbleSizes from separate cell ranges to a bubble chart with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# set XValues range for bubble chart series | How to assign bubble sizes from a column in Aspose.Cells chart | Link separate X, Y, and size ranges to a bubble chart using Aspose.Cells .NET | Create bubble chart with custom data ranges in C# Aspose.Cells example
// Tags: bubble chart series data binding Aspose.Cells C# | set XValues range Aspose.Cells chart | assign bubble sizes series Aspose.Cells | create bubble chart from cell ranges .NET | Aspose.Cells chart series configuration

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The sample creates a workbook, fills columns A, B, and C with X, Y, and size values, adds a bubble chart, adds a series using the Y range as Values, then links the XValues and BubbleSizes properties to separate cell ranges before saving the file as BubbleChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for X values, Y values, and bubble sizes
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Size");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["A4"].PutValue(3);

            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["C3"].PutValue(10);
            sheet.Cells["C4"].PutValue(15);

            // Add a bubble chart to the worksheet (rows 5‑20, columns 0‑10)
            int chartIndex = sheet.Charts.Add(ChartType.Bubble, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Add a series with Y values (Values)
            int seriesIndex = chart.NSeries.Add("Sheet1!$B$2:$B$4", true);

            // Link X values and bubble sizes to separate ranges
            chart.NSeries[seriesIndex].XValues = "Sheet1!$A$2:$A$4";
            chart.NSeries[seriesIndex].BubbleSizes = "Sheet1!$C$2:$C$4";

            // Define output file path
            string outputPath = "BubbleChart.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
