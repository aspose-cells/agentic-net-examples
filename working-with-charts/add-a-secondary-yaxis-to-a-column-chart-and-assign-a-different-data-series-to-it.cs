// Title: Aspose.Cells for .NET: Add a Secondary Y‑Axis to a Column Chart and Assign a Series (C#)
// Description: Demonstrates how to create an Excel workbook with a column chart, add two data series, plot the second series on a secondary Y‑axis, customize the secondary axis (title, min, max, major unit), and save the file using Aspose.Cells in C#.
// Keywords: Aspose.Cells secondary Y axis C# | column chart secondary axis .NET | PlotOnSecondAxis Aspose.Cells | customize secondary value axis Aspose | Excel chart multiple Y axes C# | Aspose.Cells chart example | C# add secondary axis to chart | Aspose.Cells chart NSeries
// Common Searches: Aspose.Cells add secondary Y axis to column chart | C# plot series on second axis Aspose.Cells | set secondary axis title min max Aspose.Cells | Aspose.Cells chart secondary value axis example | how to use PlotOnSecondAxis in Aspose.Cells
// Developer Intent: Create a column chart where one series uses the primary Y‑axis and another uses a secondary Y‑axis, then configure the secondary axis properties.
// Use Cases: Display sales volume and revenue together when the scales differ dramatically. | Compare a low‑volume metric (e.g., defect count) with a high‑value metric (e.g., production cost) in a single visual. | Generate financial reports that require two Y‑axes for clear data separation.
// AI Prompts: Show C# code to add a secondary Y‑axis to an Aspose.Cells column chart and assign a specific series to it. | Provide an Aspose.Cells example that sets the secondary axis title, minimum, maximum, and major unit for a column chart. | Explain how to retrieve and modify the secondary value axis object after creating a chart with Aspose.Cells in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryYAxisDemo
{
    // Demonstrates how to create an Excel workbook with a column chart, add two data series, plot the second series on a secondary Y‑axis, customize the secondary axis (title, min, max, major unit), and save the file using Aspose.Cells in C#.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add two data series
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary Y‑axis appearance
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Save the workbook
            workbook.Save("SecondaryYAxisDemo.xlsx");
        }
    }
}
