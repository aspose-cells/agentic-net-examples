// Title: Aspose.Cells C# – Add a Secondary Y‑Axis to a Column Chart and Plot a Series on It
// Description: This example creates a new workbook, inserts category data and two series, adds a column chart, assigns the second series to a secondary Y‑axis, customizes the axis title, range and major unit, and saves the file as an Excel workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | C# chart example | secondary Y axis | PlotOnSecondAxis | SecondValueAxis | column chart | Excel secondary axis | chart customization | Aspose.Cells tutorial | Excel automation
// Common Searches: Aspose.Cells add secondary Y axis to chart | C# plot series on secondary axis Aspose.Cells | set secondary axis range Aspose.Cells column chart | customize secondary value axis Aspose.Cells | Aspose.Cells PlotOnSecondAxis property example
// Developer Intent: Generate a column chart in an Excel workbook and display one data series on a secondary Y‑axis using Aspose.Cells for .NET.
// Use Cases: Display sales volume and revenue together when revenue values are an order of magnitude larger. | Compare temperature (°C) and precipitation (mm) in a single chart by assigning one metric to the secondary axis. | Create a financial dashboard that shows profit amount and profit‑margin percentage on separate Y‑axes.
// AI Prompts: Write C# code with Aspose.Cells that adds a secondary Y‑axis to a line chart and maps multiple series to it. | Explain how to configure the title, minimum, maximum, and major unit of the secondary value axis in Aspose.Cells. | Show how to hide the primary Y‑axis while keeping the secondary Y‑axis visible in an Aspose.Cells chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryYAxisDemo
{
    // This example creates a new workbook, inserts category data and two series, adds a column chart, assigns the second series to a secondary Y‑axis, customizes the axis title, range and major unit, and saves the file as an Excel workbook using Aspose.Cells for .NET.
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
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series to the chart
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
            secondaryAxis.IsVisible = true; // ensure it is shown

            // Save the workbook
            workbook.Save("ChartWithSecondaryYAxis.xlsx");
        }
    }
}
