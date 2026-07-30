// Title: Aspose.Cells C# – Add a Secondary Axis to a Combo Chart and Link Data Labels to Cells
// Description: Creates a workbook, populates month, sales, profit and label data, builds a column‑line combo chart, moves the profit series to a secondary Y‑axis, customizes the secondary axis (title, range, color) and displays custom data labels sourced from cells D2:D4.
// Keywords: Aspose.Cells combo chart secondary axis | C# Aspose.Cells cell linked data labels | secondary Y axis Aspose.Cells | chart series on second axis .NET | custom data labels from worksheet cells
// Common Searches: Aspose.Cells add secondary axis to chart C# | How to show cell‑based data labels on a secondary series in Aspose.Cells | Configure secondary value axis title and scale Aspose.Cells | Create column and line combo chart with cell labels Aspose.Cells
// Developer Intent: Generate a column‑line combo chart where the line series uses a secondary Y‑axis and its data labels are taken from worksheet cells.
// Use Cases: Financial reports that compare sales (primary axis) and profit (secondary axis) with profit values shown as custom labels. | Monthly performance dashboards that need different scales for two metrics and cell‑driven label text. | Automated workbook generation for business intelligence where axis styling and label sources are programmatically defined.
// AI Prompts: Write C# code with Aspose.Cells to add a secondary Y‑axis to a combo chart and link the line series data labels to a cell range. | Show how to set the secondary axis title, minimum, maximum, major unit, and title color in Aspose.Cells. | Explain enabling ShowCellRange and assigning LinkedSource for data labels on a secondary chart series using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsComboChartSecondaryAxis
{
    // Creates a workbook, populates month, sales, profit and label data, builds a column‑line combo chart, moves the profit series to a secondary Y‑axis, customizes the secondary axis (title, range, color) and displays custom data labels sourced from cells D2:D4.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate worksheet data ----------
            // Categories (X‑axis)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            // Primary series values (e.g., Sales)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Secondary series values (e.g., Profit)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(55);

            // Labels for the secondary series (will be shown as data labels)
            sheet.Cells["D1"].PutValue("Profit Label");
            sheet.Cells["D2"].PutValue("30k");
            sheet.Cells["D3"].PutValue("45k");
            sheet.Cells["D4"].PutValue("55k");

            // ---------- Add a combo chart ----------
            // Create a combo chart (Column + Line) placed at rows 6‑20 and columns 0‑8
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // First series – Column (primary axis)
            chart.NSeries.Add("B2:B4", true);
            // Second series – Line (secondary axis)
            chart.NSeries.Add("C2:C4", true);

            // Set category (X) data for both series
            chart.NSeries.CategoryData = "A2:A4";

            // Change the chart type of the second series to Line to create a combo effect
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // ---------- Configure the secondary Y axis ----------
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Profit (k)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 60;
            secondaryAxis.MajorUnit = 10;
            secondaryAxis.Title.Font.Color = Color.DarkGreen;

            // ---------- Show data labels for the secondary series from cells ----------
            Series secondarySeries = chart.NSeries[1];
            secondarySeries.DataLabels.ShowCellRange = true;          // Use cell range as label source
            secondarySeries.DataLabels.LinkedSource = "D2:D4";        // Cells containing the label text
            secondarySeries.DataLabels.Font.Color = Color.Blue;      // Optional styling
            secondarySeries.DataLabels.Position = LabelPositionType.OutsideEnd;

            // ---------- Optional: format primary axis ----------
            chart.ValueAxis.Title.Text = "Sales";
            chart.ValueAxis.Title.Font.Color = Color.DarkRed;

            // Save the workbook
            workbook.Save("ComboChart_SecondaryAxis_WithLabels.xlsx");
        }
    }
}
