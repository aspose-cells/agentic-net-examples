// Title: Link cell notes as data labels in a scatter chart using Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, fills columns with X values, Y values and descriptive notes, adds a scatter chart, assigns the X and Y ranges, enables data labels that pull text from the notes range (C2:C4), hides the numeric values, positions the labels above each point, applies a dark‑blue bold font, and saves the file as ScatterDataLabelsLinked.xlsx.
// Keywords: Aspose.Cells scatter chart | C# data labels linked to cells | custom scatter chart labels .NET | hide numeric values Aspose.Cells | label position above scatter points | format data label font Aspose.Cells | link notes to chart labels | Excel scatter plot annotations
// Common Searches: Aspose.Cells link data labels to cell range scatter chart | C# scatter chart custom labels from worksheet cells | how to hide values and show notes in Aspose.Cells chart | set font color and bold for scatter chart labels Aspose | position scatter chart data labels above points .NET
// Developer Intent: Display custom notes stored in worksheet cells as data labels on a scatter chart, replacing the default numeric values.
// Use Cases: Scientific plots where each point shows a descriptive comment from a worksheet column. | Sales performance scatter charts that display salesperson remarks instead of raw numbers. | Project milestone visualizations with linked notes appearing above each milestone point.
// AI Prompts: Generate C# code with Aspose.Cells to create a scatter chart and link its data labels to a range of note cells, positioning the labels above the points and hiding the numeric values. | Explain how to customize font color, size, and boldness for linked data labels in an Aspose.Cells scatter series. | Provide step‑by‑step instructions to enable cell‑range data labels for a scatter chart, hide default value labels, and apply styling.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsScatterDataLabels
{
    // This example creates a workbook, fills columns with X values, Y values and descriptive notes, adds a scatter chart, assigns the X and Y ranges, enables data labels that pull text from the notes range (C2:C4), hides the numeric values, positions the labels above each point, applies a dark‑blue bold font, and saves the file as ScatterDataLabelsLinked.xlsx.
    public class ScatterDataLabelsDemo
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate X values, Y values and descriptive notes
            // A column – X values, B column – Y values, C column – notes
            sheet.Cells["A1"].PutValue("X");
            sheet.Cells["B1"].PutValue("Y");
            sheet.Cells["C1"].PutValue("Note");

            sheet.Cells["A2"].PutValue(1);
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue("Start point");

            sheet.Cells["A3"].PutValue(2);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["C3"].PutValue("Mid point");

            sheet.Cells["A4"].PutValue(3);
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["C4"].PutValue("End point");

            // Add a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Add series using Y values; X values are set separately
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries[0].XValues = "A2:A4";

            // Enable data labels and link them to the notes range (C2:C4)
            Series series = chart.NSeries[0];
            series.DataLabels.ShowCellRange = true;          // Use cell range for label text
            series.DataLabels.LinkedSource = "C2:C4";        // Range containing descriptive notes
            series.DataLabels.ShowValue = false;            // Hide numeric value, show only notes
            series.DataLabels.Position = LabelPositionType.Above; // Position labels above points

            // Optional: style the data labels for better visibility
            series.DataLabels.Font.Color = Color.DarkBlue;
            series.DataLabels.Font.IsBold = true;

            // Save the workbook
            workbook.Save("ScatterDataLabelsLinked.xlsx");
        }
    }
}
