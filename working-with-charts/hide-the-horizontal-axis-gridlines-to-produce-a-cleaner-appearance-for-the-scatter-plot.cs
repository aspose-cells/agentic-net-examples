// Title: Hide horizontal (category) axis gridlines in a C# Aspose.Cells scatter chart
// AI Prompts: Generate C# code with Aspose.Cells that creates a scatter chart and sets CategoryAxis.MajorGridLines.IsVisible to false. | Show how to disable both major and minor horizontal gridlines for a scatter plot in Aspose.Cells and then save the workbook. | Provide a step‑by‑step example that builds X/Y sample data, adds a scatter chart, and hides its category axis gridlines using Aspose.Cells.
// Common Searches: aspocells c# scatter chart hide category axis gridlines | remove horizontal gridlines from Excel scatter plot using Aspose.Cells | how to turn off major gridlines on scatter chart axis in C# | Aspose.Cells hide minor gridlines on chart category axis | C# create scatter chart without horizontal gridlines Aspose.Cells
// Tags: Aspose.Cells scatter chart hide category axis gridlines | C# chart axis major gridlines visibility | Aspose.Cells disable minor gridlines on chart | Excel scatter plot gridline removal with Aspose.Cells | CategoryAxis gridlines IsVisible false Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsScatterGridlinesDemo
{
    // Creates a workbook, adds X/Y data, generates a scatter chart, hides both major and minor horizontal (category) axis gridlines, and saves the file as ScatterChart_NoHorizontalGridlines.xlsx.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a scatter chart (X values in column A, Y values in column B)
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(2);
            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(4);
            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(6);
            sheet.Cells["A5"].PutValue(4);
            sheet.Cells["B5"].PutValue(8);
            sheet.Cells["A6"].PutValue(5);
            sheet.Cells["B6"].PutValue(10);

            // Add a scatter chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the scatter series (X values, Y values)
            chart.NSeries.Add("B2:B6", true);          // Y values
            chart.NSeries[0].XValues = "A2:A6";        // X values

            // Hide horizontal (category) axis gridlines for a cleaner look
            // Major gridlines
            chart.CategoryAxis.MajorGridLines.IsVisible = false;
            // Minor gridlines (optional, also hide)
            chart.CategoryAxis.MinorGridLines.IsVisible = false;

            // Save the workbook to a file
            workbook.Save("ScatterChart_NoHorizontalGridlines.xlsx");
        }
    }
}
