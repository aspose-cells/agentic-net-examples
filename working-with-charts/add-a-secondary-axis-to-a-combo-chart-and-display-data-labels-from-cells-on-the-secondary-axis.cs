// Title: Create a column‑line combo chart with a secondary axis and cell‑linked data labels using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to create a column‑line combo chart, set the second series as a line plotted on the secondary vertical axis, and link its data labels to a range of worksheet cells. | Demonstrate configuring the secondary value axis range, title, and custom font styling for data labels sourced from cells in an Aspose.Cells chart.
// Common Searches: Aspose.Cells C# create combo chart with secondary Y axis and cell‑based data labels | link data labels to cell range in secondary axis line series Aspose.Cells | set secondary value axis properties for line series in Aspose.Cells chart | column and line series combo chart with custom label formatting using Aspose.Cells .NET | how to plot profit series on secondary axis in Excel using Aspose.Cells C#
// Tags: combined column line chart secondary axis Aspose.Cells C# | line series cell‑linked data labels Aspose.Cells | secondary axis scaling Aspose.Cells | cell range label formatting Aspose.Cells | column and line series chart Aspose.Cells .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsComboChartSecondaryAxis
{
    // The example creates a workbook, adds month, sales, profit, and label data, builds a column‑line combo chart, moves the profit series to a secondary vertical axis, configures the secondary axis range and title, links the profit series' data labels to cells D2:D4 with custom font styling, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Populate sample data
            // -------------------------------------------------
            // Primary categories
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Primary series (Column)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Secondary series (Line) – values will be shown on secondary axis
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);

            // Data labels for secondary series (taken from cells D2:D4)
            sheet.Cells["D1"].PutValue("Profit Label");
            sheet.Cells["D2"].PutValue("30k");
            sheet.Cells["D3"].PutValue("45k");
            sheet.Cells["D4"].PutValue("55k");

            // -------------------------------------------------
            // Add a combo chart (Column + Line)
            // -------------------------------------------------
            // Add a chart of type Column; later we will change the second series to Line
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series – Column (primary axis)
            chart.NSeries.Add("B2:B4", true);
            // Second series – initially Column, will be changed to Line
            chart.NSeries.Add("C2:C4", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Configure the second series as Line and plot on secondary axis
            // -------------------------------------------------
            Series secondarySeries = chart.NSeries[1];
            secondarySeries.Type = ChartType.Line;               // Change series type to Line
            secondarySeries.PlotOnSecondAxis = true;             // Plot on secondary Y axis

            // -------------------------------------------------
            // Customize secondary value axis (optional)
            // -------------------------------------------------
            Axis secValueAxis = chart.SecondValueAxis;
            secValueAxis.Title.Text = "Profit (k)";
            secValueAxis.MinValue = 0;
            secValueAxis.MaxValue = 60;
            secValueAxis.MajorUnit = 10;

            // -------------------------------------------------
            // Show data labels for the secondary series using cell range
            // -------------------------------------------------
            secondarySeries.DataLabels.ShowCellRange = true;     // Use cell range for labels
            secondarySeries.DataLabels.LinkedSource = "D2:D4";   // Cells containing label text
            secondarySeries.DataLabels.Font.Color = Color.DarkGreen;
            secondarySeries.DataLabels.Font.IsBold = true;
            secondarySeries.DataLabels.Position = LabelPositionType.OutsideEnd;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ComboChart_SecondaryAxis_WithCellLabels.xlsx");
        }
    }
}
