// Title: Set custom fill and border colors for each series in a stacked bar chart with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that builds a stacked bar chart using Aspose.Cells and assigns a distinct RGB fill color to each series via the Series.Area.ForegroundColor property. | Show how to apply a black border to every series in an Aspose.Cells stacked bar chart while preserving custom fill colors. | Create a complete Aspose.Cells example that populates data, adds a stacked bar chart, and customizes both series fill and border colors before saving the workbook.
// Common Searches: c# aspocells set individual series fill colors in a stacked bar chart | aspocells change series area foregroundcolor for stacked bar chart .net | example of custom border color for each series in Aspose.Cells chart | how to assign different RGB colors to series in an Aspose.Cells stacked bar chart
// Tags: Aspose.Cells series area foregroundcolor | stacked bar chart custom series colors .NET | C# Aspose.Cells set series border color | Aspose.Cells chart series fill color RGB | customize stacked bar series appearance Aspose.Cells

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSeriesColorDemo
{
    // Demonstrates creating a workbook, adding sample data, inserting a stacked bar chart, and using the Series.Area.ForegroundColor and Series.Border.Color properties to apply unique fill and border colors to each series before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a stacked bar chart
            // Categories
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");

            // Series 1
            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Series 2
            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Series 3
            sheet.Cells["D1"].PutValue("Series3");
            sheet.Cells["D2"].PutValue(12);
            sheet.Cells["D3"].PutValue(22);
            sheet.Cells["D4"].PutValue(32);

            // Add a stacked bar chart
            int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (B2:D4) and categories (A2:A4)
            chart.NSeries.Add("B2:D4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Assign custom colors to each series using the Area.ForegroundColor property
            // (Series.Style is not a public member; using Area.ForegroundColor achieves the same effect)
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189);   // Blueish
            chart.NSeries[1].Area.ForegroundColor = Color.FromArgb(192, 80, 77);   // Reddish
            chart.NSeries[2].Area.ForegroundColor = Color.FromArgb(155, 187, 89);  // Greenish

            // Optionally, set border colors to emphasize the series
            chart.NSeries[0].Border.Color = Color.Black;
            chart.NSeries[1].Border.Color = Color.Black;
            chart.NSeries[2].Border.Color = Color.Black;

            // Save the workbook
            workbook.Save("StackedBarCustomSeriesColors.xlsx");
        }
    }
}
