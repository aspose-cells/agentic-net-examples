using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSecondaryAxisDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["A4"].PutValue("Q3");
            sheet.Cells["A5"].PutValue("Q4");

            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(150);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(250);
            sheet.Cells["B5"].PutValue(300);

            sheet.Cells["C1"].PutValue("Revenue");
            sheet.Cells["C2"].PutValue(1200);
            sheet.Cells["C3"].PutValue(1800);
            sheet.Cells["C4"].PutValue(2400);
            sheet.Cells["C5"].PutValue(3000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // First series (primary axis)
            chart.NSeries.Add("B2:B5", true);
            // Second series (secondary axis)
            chart.NSeries.Add("C2:C5", true);
            chart.NSeries[1].PlotOnSecondAxis = true; // Enable secondary axis for the second series

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Configure secondary value axis title
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.Title.Text = "Revenue (USD)";
            secondaryAxis.Title.IsVisible = true;
            secondaryAxis.Title.Font.IsBold = true;
            secondaryAxis.Title.Font.IsItalic = true;

            // Optional: make primary axis titles visible for context
            chart.ValueAxis.Title.Text = "Sales";
            chart.ValueAxis.Title.IsVisible = true;
            chart.CategoryAxis.Title.Text = "Quarter";
            chart.CategoryAxis.Title.IsVisible = true;

            // Save the workbook
            workbook.Save("SecondaryAxisTitleDemo.xlsx");
        }
    }
}