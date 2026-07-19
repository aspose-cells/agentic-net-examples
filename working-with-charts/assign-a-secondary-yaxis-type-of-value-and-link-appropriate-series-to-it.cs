// Title: How to Assign a Secondary Value Y‑Axis to a Series in Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, defines two series, plots the second series on a secondary value Y‑axis, customizes axis titles, range and major unit, and saves the file as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells secondary Y axis C# | dual axis chart Aspose.Cells | plot series on secondary axis .NET | configure secondary value axis Aspose.Cells | secondary axis range Aspose.Cells
// Common Searches: Aspose.Cells plot series on secondary Y axis | set secondary axis minimum maximum Aspose.Cells C# | dual axis column chart example Aspose.Cells | secondary value axis title Aspose.Cells | how to use PlotOnSecondAxis Aspose.Cells
// Developer Intent: Generate a chart, assign a series to a secondary value Y‑axis, and configure that axis’s properties programmatically.
// Use Cases: Show sales volume and revenue together, using revenue on a secondary axis with a larger scale. | Compare temperature and precipitation in a weather report, placing precipitation on a secondary axis. | Display budgeted versus actual costs, with actual costs plotted on a secondary axis for clearer variance analysis.
// AI Prompts: Write C# code with Aspose.Cells that adds a line series to the secondary Y‑axis and applies a custom number format. | Provide a step‑by‑step guide to customize tick marks, label rotation, and gridlines of a secondary value axis in an Aspose.Cells chart. | Explain how to automatically select which series should use the secondary axis based on data magnitude using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryYAxisDemo
{
    // Creates a workbook, adds sample data, inserts a column chart, defines two series, plots the second series on a secondary value Y‑axis, customizes axis titles, range and major unit, and saves the file as an XLSX document using Aspose.Cells for .NET.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Primary Series");
            worksheet.Cells["B2"].PutValue(100);
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["B4"].PutValue(300);

            worksheet.Cells["C1"].PutValue("Secondary Series");
            worksheet.Cells["C2"].PutValue(5000);
            worksheet.Cells["C3"].PutValue(3000);
            worksheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Add two series: first uses primary Y‑axis, second will use secondary Y‑axis
            chart.NSeries.Add("B2:B4", true); // Primary series
            chart.NSeries.Add("C2:C4", true); // Secondary series
            chart.NSeries.CategoryData = "A2:A4";

            // Link the second series to the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Configure the secondary Y‑axis (which is of type Value by default)
            Axis secondaryValueAxis = chart.SecondValueAxis;
            secondaryValueAxis.Title.Text = "Secondary Axis (Value)";
            secondaryValueAxis.MinValue = 0;
            secondaryValueAxis.MaxValue = 6000;
            secondaryValueAxis.MajorUnit = 1000;

            // Optionally customize the primary Y‑axis
            chart.ValueAxis.Title.Text = "Primary Axis (Value)";

            // Save the workbook
            workbook.Save("SecondaryYAxisDemo.xlsx");
        }
    }
}
