// Title: Aspose.Cells C# – Combo chart with secondary Y‑axis and cell‑linked data labels
// Description: Demonstrates how to create a workbook, add a column‑line combo chart, plot the line series on a secondary Y‑axis, configure axis limits and title, and display data labels that are linked to a worksheet cell range with custom font styling, then save the file.
// Keywords: Aspose.Cells combo chart secondary axis | C# secondary Y axis chart | cell linked data labels Aspose.Cells | Aspose.Cells line series on secondary axis | custom axis limits Aspose.Cells | Aspose.Cells chart formatting | Excel combo chart Aspose.Cells .NET
// Common Searches: Aspose.Cells add secondary axis to combo chart C# | show data labels from cells on secondary series Aspose.Cells | configure secondary Y axis range and title Aspose.Cells | convert column chart to combo chart Aspose.Cells .NET | link chart data labels to worksheet cells Aspose.Cells
// Developer Intent: Create a combo chart with a secondary Y‑axis and bind the secondary series' data labels to cells.
// Use Cases: Build a column‑line combo chart where the line series uses a different scale (e.g., units vs. quantity). | Display custom text such as "5k units" from a worksheet column next to each point of the secondary series. | Set secondary axis title, minimum, maximum, and major unit values while applying bold dark‑blue font to the linked labels.
// AI Prompts: Generate C# Aspose.Cells code that adds a secondary Y‑axis to a combo chart and links the secondary series' data labels to a specified cell range. | Show how to configure axis titles, min/max values, and custom font styling for data labels on a secondary axis in Aspose.Cells. | Provide an example of converting a column chart to a combo chart by changing the second series to a line type and plotting it on the secondary axis using Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartExample
{
    // Demonstrates how to create a workbook, add a column‑line combo chart, plot the line series on a secondary Y‑axis, configure axis limits and title, and display data labels that are linked to a worksheet cell range with custom font styling, then save the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // ---------- Populate worksheet data ----------
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            // Primary series values (plotted on primary Y axis)
            sheet.Cells["B1"].PutValue("Primary Series");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);

            // Secondary series values (plotted on secondary Y axis)
            sheet.Cells["C1"].PutValue("Secondary Series");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Labels for secondary series (will be shown as data labels)
            sheet.Cells["D1"].PutValue("Sec Labels");
            sheet.Cells["D2"].PutValue("5k units");
            sheet.Cells["D3"].PutValue("3k units");
            sheet.Cells["D4"].PutValue("1k units");

            // ---------- Add a combo chart ----------
            // Create a column chart; later we will change the second series to a line type
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add primary series (column)
            chart.NSeries.Add("B2:B4", true);
            // Add secondary series (line)
            chart.NSeries.Add("C2:C4", true);
            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A4";

            // Change the second series to a line type to create a combo effect
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // ---------- Configure secondary Y axis ----------
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis (Units)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // ---------- Show data labels for the secondary series ----------
            // Enable data labels and link them to the range D2:D4
            Series secondarySeries = chart.NSeries[1];
            secondarySeries.DataLabels.ShowCellRange = true;
            secondarySeries.DataLabels.LinkedSource = "D2:D4";
            // Optional: customize label appearance
            secondarySeries.DataLabels.Font.Color = Color.DarkBlue;
            secondarySeries.DataLabels.Font.IsBold = true;

            // Save the workbook with the chart
            workbook.Save("ComboChart_With_SecondaryAxis.xlsx");
        }
    }
}
