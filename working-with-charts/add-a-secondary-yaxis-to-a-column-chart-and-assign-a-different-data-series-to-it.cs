// Title: Create a column chart with a secondary Y‑axis in Aspose.Cells for .NET (C#)
// Description: Shows how to build an Excel workbook, add a column chart with two data series, plot the second series on a secondary Y‑axis, customize the secondary axis title and scale, and save the file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells secondary Y axis C# | column chart secondary axis Aspose.Cells | plot series on second axis .NET | customize secondary value axis Aspose | Aspose.Cells chart axis scaling | C# Excel chart secondary axis example | NSeries PlotOnSecondAxis Aspose
// Common Searches: Aspose.Cells secondary Y axis column chart C# | How to add a secondary value axis in Aspose.Cells | C# plot series on second axis Excel chart Aspose | Set secondary axis title and range Aspose.Cells | Example of secondary axis with column chart Aspose
// Developer Intent: Generate an Excel column chart where one series uses the primary Y‑axis and another series uses a secondary Y‑axis, with optional formatting of the secondary axis.
// Use Cases: Display units sold (primary axis) alongside revenue (secondary axis) in a single chart. | Compare a small‑scale metric such as defect count with a large‑scale metric like production cost. | Create financial reports that need separate scaling for quantity and monetary values. | Build dashboards where temperature (primary) and humidity (secondary) are visualized together.
// AI Prompts: Write C# code with Aspose.Cells that adds a secondary Y‑axis to a bar chart and assigns the third data series to it. | Explain how to set the secondary axis title, minimum, maximum, and major unit values for an Aspose.Cells chart. | Provide step‑by‑step instructions to retrieve and modify the secondary axis after a chart has been created in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryYAxisDemo
{
    // Shows how to build an Excel workbook, add a column chart with two data series, plot the second series on a secondary Y‑axis, customize the secondary axis title and scale, and save the file using Aspose.Cells for .NET.
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
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(4000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series: first uses primary Y‑axis, second will use secondary Y‑axis
            chart.NSeries.Add("B2:B4", true); // Series 1
            chart.NSeries.Add("C2:C4", true); // Series 2
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize the secondary Y‑axis appearance
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis (Units)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // Save the workbook
            workbook.Save("ColumnChartWithSecondaryYAxis.xlsx");
        }
    }
}
