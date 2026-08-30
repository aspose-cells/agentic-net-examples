// Title: Create a combined column and line chart with primary and secondary axes in Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that uses Aspose.Cells to build a column chart for sales data and adds a profit series as a line plotted on the secondary value axis. | Show how to change a specific series type to Line and enable PlotOnSecondAxis in an Aspose.Cells chart. | Provide a complete Aspose.Cells example that saves the mixed column‑line chart to an Excel workbook.
// Common Searches: aspocells mixed column line chart secondary axis c# example | how to set a series to secondary axis in Aspose.Cells chart | C# Aspose.Cells create combined chart with column and line series | Aspose.Cells plot profit as line on secondary axis | changing series type to line in Aspose.Cells chart programmatically
// Tags: combined column line chart Aspose.Cells | plot series on secondary axis Aspose.Cells | set series type line Aspose.Cells | mixed chart creation Aspose.Cells .NET | customize secondary value axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace CombinedColumnLineChartDemo
{
    // Demonstrates creating a workbook, inserting sales and profit data, adding a column chart, converting the profit series to a line, assigning it to the secondary axis, customizing the secondary axis title, and saving the result as CombinedColumnLineChart.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            // Category labels
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["A5"].PutValue("Apr");

            // Column series values (e.g., Sales)
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Line series values (e.g., Profit)
            sheet.Cells["C1"].PutValue("Profit");
            sheet.Cells["C2"].PutValue(30);
            sheet.Cells["C3"].PutValue(45);
            sheet.Cells["C4"].PutValue(50);
            sheet.Cells["C5"].PutValue(70);

            // Add a chart of type Column (primary axis will host column series)
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Add the column series (primary axis)
            chart.NSeries.Add("B2:B5", true);          // true => use categories from column A
            // Add the line series (will be moved to secondary axis)
            chart.NSeries.Add("C2:C5", true);

            // Set category (X) axis data
            chart.NSeries.CategoryData = "A2:A5";

            // Change the second series to a line type
            chart.NSeries[1].Type = ChartType.Line;

            // Plot the second series on the secondary value axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Optional: customize secondary axis title
            chart.SecondValueAxis.Title.Text = "Profit (Secondary Axis)";

            // Save the workbook
            workbook.Save("CombinedColumnLineChart.xlsx");
        }
    }
}
