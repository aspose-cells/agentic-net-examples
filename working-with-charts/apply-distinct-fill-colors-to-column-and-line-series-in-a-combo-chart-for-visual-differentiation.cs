using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsComboChartDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category labels
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column series data (e.g., Sales)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Line series data (e.g., Profit)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(50);
            sheet.Cells["C5"].PutValue(70);

            // Add a combo chart (Column + Line)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the two series to the chart
            chart.NSeries.Add("B2:B5", true); // Column series
            chart.NSeries.Add("C2:C5", true); // Line series

            // Set category (X‑axis) data
            chart.NSeries.CategoryData = "A2:A5";

            // Convert the second series to a line type
            chart.NSeries[1].Type = ChartType.Line;

            // Apply distinct fill colors
            // Column series – set solid fill color via Area.ForegroundColor
            chart.NSeries[0].Area.ForegroundColor = Color.FromArgb(79, 129, 189); // a blue shade

            // Line series – set line color via SeriesLines.Color
            chart.NSeries[1].SeriesLines.Color = Color.FromArgb(192, 80, 77); // a red shade

            // Optional: give the chart a title
            chart.Title.Text = "Sales (Column) vs Profit (Line)";

            // Save the workbook
            workbook.Save("ComboChartDistinctColors.xlsx");
        }
    }
}