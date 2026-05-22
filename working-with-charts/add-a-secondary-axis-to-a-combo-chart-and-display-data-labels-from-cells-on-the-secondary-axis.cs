using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartSecondaryAxis
{
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
            // Categories (X axis)
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

            // Data label texts for the secondary series (taken from cells)
            sheet.Cells["D1"].PutValue("Label");
            sheet.Cells["D2"].PutValue("High");
            sheet.Cells["D3"].PutValue("Medium");
            sheet.Cells["D4"].PutValue("Low");

            // -------------------------------------------------
            // Add a combo chart (Column + Line) to the worksheet
            // -------------------------------------------------
            // The chart type is Column; later we can change the second series to Line if desired.
            int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the primary series (first series)
            chart.NSeries.Add("B2:B4", true);
            // Add the secondary series (second series)
            chart.NSeries.Add("C2:C4", true);

            // Set the category (X‑axis) data for both series
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // Configure the secondary axis
            // -------------------------------------------------
            // Plot the second series on the secondary Y axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: give the secondary axis a title and adjust its scale
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Secondary Axis (Units)";
            secondaryAxis.MinValue = 0;
            secondaryAxis.MaxValue = 6000;
            secondaryAxis.MajorUnit = 1000;

            // -------------------------------------------------
            // Show data labels for the secondary series from cells
            // -------------------------------------------------
            Series secondarySeries = chart.NSeries[1];
            secondarySeries.DataLabels.ShowCellRange = true;          // Use cell range for labels
            secondarySeries.DataLabels.LinkedSource = "D2:D4";        // Cells containing label texts
            secondarySeries.DataLabels.ShowValue = false;            // Hide the numeric value if not needed
            secondarySeries.DataLabels.Position = LabelPositionType.OutsideEnd;

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("ComboChart_SecondaryAxis_WithCellLabels.xlsx");
        }
    }
}