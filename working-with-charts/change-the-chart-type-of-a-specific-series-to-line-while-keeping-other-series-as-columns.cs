// Title: Change a single series to a line chart while keeping other series as columns using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that sets the second NSeries of a clustered column chart to ChartType.Line and leaves the first series as ChartType.Column in Aspose.Cells. | Show how to customize the line series border color and weight after changing its chart type with Aspose.Cells. | Explain the steps to modify only one series type in a mixed chart without affecting the other series in a .NET workbook.
// Common Searches: aspnet change chart series type to line only second series Aspose.Cells | mixed column and line chart Aspose.Cells C# example | how to set individual series chart type in Aspose.Cells workbook | customize line series appearance Aspose.Cells chart C#
// Tags: set chart type for specific series Aspose.Cells | mixed column‑line chart C# Aspose.Cells | line series border color Aspose.Cells | NSeries type assignment Aspose.Cells | chart series type conversion .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsSeriesTypeDemo
{
    // The example creates a workbook, adds sample sales and profit data, inserts a clustered column chart covering both series, then changes the second series (Profit) to a line chart, customizes its border color to red and line weight, and saves the workbook as SeriesTypeChanged.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category column
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // First series (will stay as Column)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Second series (will be changed to Line)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(60);

            // Add a clustered column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for both series
            chart.NSeries.Add("B2:C4", true);          // Values for both series
            chart.NSeries.CategoryData = "A2:A4";      // Categories

            // Change the type of the second series (index 1) to Line
            chart.NSeries[1].Type = ChartType.Line;

            // Optional: customize appearance of the line series
            chart.NSeries[1].Border.Color = Color.Red;
            chart.NSeries[1].Border.Weight = WeightType.MediumLine;

            // Save the workbook
            workbook.Save("SeriesTypeChanged.xlsx");
        }
    }
}
