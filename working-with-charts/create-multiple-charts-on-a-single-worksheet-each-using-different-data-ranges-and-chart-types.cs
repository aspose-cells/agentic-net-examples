// Title: Add Column, Pie, and Line Charts to a Single Worksheet with Aspose.Cells for .NET (C#)
// Description: This example creates a new workbook, populates three distinct data ranges (A1:B5, C1:D5, E1:F5), and inserts a column chart, a pie chart, and a line chart on the same sheet. Each chart is linked to its own range, positioned at different cell locations, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells multiple charts C# | add column chart Aspose.Cells | pie chart Aspose.Cells .NET | line chart worksheet Aspose | chart positioning Aspose.Cells | set chart data range programmatically | C# Excel chart automation | Aspose.Cells chart examples
// Common Searches: how to add several chart types to one worksheet using Aspose.Cells | Aspose.Cells C# create column, pie and line charts | set chart data range and location with Aspose.Cells | multiple charts on a single Excel sheet Aspose | Aspose.Cells add chart programmatically
// Developer Intent: Insert three different chart types on one worksheet, each bound to its own data range and placed at a specific location.
// Use Cases: Build a sales dashboard that shows category totals (column), product mix (pie), and monthly trends (line) in a single Excel file. | Generate automated reports where pre‑configured visualizations are ready for dynamic data insertion. | Create a template workbook with embedded charts that can be reused across multiple projects.
// AI Prompts: Show how to customize the titles, legends, and axis labels for each chart in the sample. | Provide code to export the three charts as separate PNG images while keeping them in the workbook. | Explain how to bind the charts to named ranges so they update automatically when data changes.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleChartsDemo
{
    // This example creates a new workbook, populates three distinct data ranges (A1:B5, C1:D5, E1:F5), and inserts a column chart, a pie chart, and a line chart on the same sheet. Each chart is linked to its own range, positioned at different cell locations, and the workbook is saved as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -----------------------------
            // Populate data for three charts
            // -----------------------------

            // Data for Chart 1 (Column Chart) - Range A1:B5
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat{i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10); // 20,30,40,50
            }

            // Data for Chart 2 (Pie Chart) - Range C1:D5
            sheet.Cells["C1"].PutValue("Item");
            sheet.Cells["D1"].PutValue("Amount");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"C{i}"].PutValue($"Item{i - 1}");
                sheet.Cells[$"D{i}"].PutValue(i * 15); // 30,45,60,75
            }

            // Data for Chart 3 (Line Chart) - Range E1:F5
            sheet.Cells["E1"].PutValue("Month");
            sheet.Cells["F1"].PutValue("Sales");
            for (int i = 2; i <= 5; i++)
            {
                sheet.Cells[$"E{i}"].PutValue($"M{i - 1}");
                sheet.Cells[$"F{i}"].PutValue(i * 20); // 40,60,80,100
            }

            // ---------------------------------
            // Add Chart 1: Column Chart (A1:B5)
            // ---------------------------------
            int chartIndex1 = sheet.Charts.Add(
                ChartType.Column,          // Chart type
                "A1:B5",                  // Data range
                true,                     // Plot by column (vertical)
                7, 0,                     // Top row, Left column (position on sheet)
                20, 5);                   // Bottom row, Right column (size)

            Chart chart1 = sheet.Charts[chartIndex1];
            chart1.Title.Text = "Column Chart Example";

            // ---------------------------------
            // Add Chart 2: Pie Chart (C1:D5)
            // ---------------------------------
            int chartIndex2 = sheet.Charts.Add(
                ChartType.Pie,
                "C1:D5",
                true,
                7, 7,
                20, 12);

            Chart chart2 = sheet.Charts[chartIndex2];
            chart2.Title.Text = "Pie Chart Example";

            // ---------------------------------
            // Add Chart 3: Line Chart (E1:F5)
            // ---------------------------------
            int chartIndex3 = sheet.Charts.Add(
                ChartType.Line,
                "E1:F5",
                true,
                22, 0,
                35, 5);

            Chart chart3 = sheet.Charts[chartIndex3];
            chart3.Title.Text = "Line Chart Example";

            // Save the workbook to a file
            workbook.Save("MultipleChartsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
