using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsPieChartLabelFromCells
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["A4"].PutValue("Cherry");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(30);
            sheet.Cells["B3"].PutValue(45);
            sheet.Cells["B4"].PutValue(25);

            // Cells that contain the label text we want to display
            sheet.Cells["C1"].PutValue("Label");
            sheet.Cells["C2"].PutValue("Red Fruit");
            sheet.Cells["C3"].PutValue("Yellow Fruit");
            sheet.Cells["C4"].PutValue("Red Fruit");

            // Add a pie chart
            int chartIndex = sheet.Charts.Add(ChartType.Pie, 6, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series (values)
            chart.NSeries.Add("B2:B4", true);
            // Set the category (slice names)
            chart.NSeries.CategoryData = "A2:A4";

            // Access the first (and only) series
            Series series = chart.NSeries[0];

            // Enable data labels and link them to the cells in column C
            series.DataLabels.ShowValue = true;               // Show the value (required to make labels visible)
            series.DataLabels.LinkedSource = "C2:C4";          // Link label text to cells
            series.DataLabels.NumberFormatLinked = true;      // Keep number format linked to source cells
            series.DataLabels.IsNeverOverlap = true;          // Optional: avoid overlapping labels for pie chart

            // Save the workbook
            workbook.Save("PieChartLabelFromCells.xlsx");
        }
    }
}