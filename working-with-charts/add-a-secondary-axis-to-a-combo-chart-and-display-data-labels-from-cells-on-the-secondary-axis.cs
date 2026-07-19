// Title: Aspose.Cells for .NET – Create a Column‑Line Combo Chart with a Secondary Axis and Cell‑Based Data Labels (C#)
// Description: This example builds a new workbook, fills it with categories, revenue and growth data, adds a column‑line combo chart, plots the growth series on a secondary Y‑axis, customizes the axis range and title, links data labels to cells in column D, hides default values, and saves the file as an Excel workbook.
// Keywords: Aspose.Cells C# | combo chart secondary axis | cell linked data labels | column line chart Aspose.Cells | customize secondary value axis | Excel chart automation .NET | financial dashboard chart | Aspose.Cells chart example
// Common Searches: Aspose.Cells add secondary Y axis to combo chart | display data labels from worksheet cells in Aspose.Cells | C# create column‑line combo chart with secondary axis | set secondary axis title and range Aspose.Cells | bind chart data labels to cell range Aspose.Cells .NET
// Developer Intent: Generate a combo chart with a secondary axis and show custom labels taken from worksheet cells.
// Use Cases: Financial reports that show revenue columns and growth percentages on a separate axis with percentage labels from cells. | Sales dashboards where a line series (e.g., profit margin) needs its own axis and cell‑based annotations. | Automated Excel generation for presentations that require mixed chart types and precise axis scaling.
// AI Prompts: Write C# code using Aspose.Cells to create a column‑line combo chart, plot the second series on a secondary Y‑axis, set axis limits and title, and link its data labels to a cell range. | Show how to hide default data label values and display custom text from worksheet cells for a secondary‑axis series in Aspose.Cells .NET. | Explain the steps to customize the secondary value axis (min, max, major unit) and bind data labels to cells in an Aspose.Cells combo chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartSecondaryAxis
{
    // This example builds a new workbook, fills it with categories, revenue and growth data, adds a column‑line combo chart, plots the growth series on a secondary Y‑axis, customizes the axis range and title, links data labels to cells in column D, hides default values, and saves the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate sample data ----------
            // Categories (X‑axis)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            // Primary series values (displayed on primary Y‑axis)
            sheet.Cells["B1"].PutValue("Revenue");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);
            sheet.Cells["B5"].PutValue(170);

            // Secondary series values (displayed on secondary Y‑axis)
            sheet.Cells["C1"].PutValue("Growth %");
            sheet.Cells["C2"].PutValue(5);
            sheet.Cells["C3"].PutValue(8);
            sheet.Cells["C4"].PutValue(6);
            sheet.Cells["C5"].PutValue(9);

            // Labels for the secondary series (will be shown as data labels)
            sheet.Cells["D1"].PutValue("GrowthLabel");
            sheet.Cells["D2"].PutValue("5%");
            sheet.Cells["D3"].PutValue("8%");
            sheet.Cells["D4"].PutValue("6%");
            sheet.Cells["D5"].PutValue("9%");

            // ---------- Add a combo chart ----------
            // Column chart for the primary series, line chart for the secondary series
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary series (Revenue)
            chart.NSeries.Add("B2:B5", true);
            // Add secondary series (Growth %)
            chart.NSeries.Add("C2:C5", true);
            // Set category (X) data
            chart.NSeries.CategoryData = "A2:A5";

            // Change the second series to a line type to create a combo chart
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the second series on the secondary Y‑axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // ---------- Customize the secondary axis ----------
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Growth Percentage";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 10;
            secondaryAxis.MajorUnit = 2;

            // ---------- Show data labels for the secondary series ----------
            Series secondarySeries = chart.NSeries[1];
            // Enable data labels and bind them to the cells in column D
            secondarySeries.DataLabels.ShowCellRange = true;
            secondarySeries.DataLabels.LinkedSource = "D2:D5";
            // Optional: hide the default value and show only the linked text
            secondarySeries.DataLabels.ShowValue = false;
            secondarySeries.DataLabels.ShowCategoryName = false;
            secondarySeries.DataLabels.ShowSeriesName = false;

            // ---------- Save the workbook ----------
            workbook.Save("ComboChart_SecondaryAxis.xlsx");
        }
    }
}
