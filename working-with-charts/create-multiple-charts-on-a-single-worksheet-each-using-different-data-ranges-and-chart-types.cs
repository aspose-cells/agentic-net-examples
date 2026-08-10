// Title: Add Multiple Different Charts to a Single Worksheet with Aspose.Cells for .NET (C#)
// Description: This C# example demonstrates how to create a new workbook, populate three separate data tables, and insert a column chart, a pie chart, and a line‑with‑data‑markers chart on the same worksheet. Each chart is assigned its own data range via SetChartDataRange, positioned in distinct cell blocks, given a title, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells multiple charts | C# add column chart | Aspose.Cells pie chart example | line chart with data markers Aspose | SetChartDataRange .NET | position charts worksheet Aspose.Cells | save workbook with charts | Aspose.Cells chart types
// Common Searches: how to add several chart types to one sheet using Aspose.Cells | Aspose.Cells set separate data ranges for multiple charts | position multiple charts on the same worksheet .NET | create column, pie, and line charts with Aspose.Cells | Aspose.Cells C# example multiple charts
// Developer Intent: Insert several charts of different types into a single worksheet, each with its own data source and layout, using Aspose.Cells for .NET.
// Use Cases: Generate a sales dashboard workbook that shows product categories (column), market share (pie), and monthly trends (line) side‑by‑side. | Automate financial reporting where each key metric is visualized with a distinct chart placed in a predefined cell range. | Create an Excel export for a BI tool that includes multiple visualizations without manual user interaction.
// AI Prompts: Write C# code with Aspose.Cells to add a bar, doughnut, and scatter chart to the same worksheet, each using a unique data range and positioned in separate cell blocks. | Show how to update the data source of an existing chart in an Aspose.Cells workbook without recreating the chart. | Provide a step‑by‑step guide to programmatically resize and style multiple charts on one sheet using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleChartsDemo
{
    // This C# example demonstrates how to create a new workbook, populate three separate data tables, and insert a column chart, a pie chart, and a line‑with‑data‑markers chart on the same worksheet. Each chart is assigned its own data range via SetChartDataRange, positioned in distinct cell blocks, given a title, and the workbook is saved as an XLSX file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // Prepare data for the first chart (Column Chart)
            // -------------------------------------------------
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Add the first chart (Column) and position it on the sheet
            int colChartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 5);
            Chart colChart = sheet.Charts[colChartIndex];
            // Set the data range for the column chart (vertical series)
            colChart.SetChartDataRange("A1:B5", true);
            colChart.Title.Text = "Column Chart Example";

            // -------------------------------------------------
            // Prepare data for the second chart (Pie Chart)
            // -------------------------------------------------
            sheet.Cells["D1"].PutValue("Item");
            sheet.Cells["E1"].PutValue("Amount");
            sheet.Cells["D2"].PutValue("X");
            sheet.Cells["E2"].PutValue(25);
            sheet.Cells["D3"].PutValue("Y");
            sheet.Cells["E3"].PutValue(35);
            sheet.Cells["D4"].PutValue("Z");
            sheet.Cells["E4"].PutValue(40);

            // Add the second chart (Pie) and position it on the sheet
            int pieChartIndex = sheet.Charts.Add(ChartType.Pie, 6, 7, 20, 12);
            Chart pieChart = sheet.Charts[pieChartIndex];
            // Set the data range for the pie chart (vertical series)
            pieChart.SetChartDataRange("D1:E4", true);
            pieChart.Title.Text = "Pie Chart Example";

            // -------------------------------------------------
            // Prepare data for the third chart (Line with Data Markers)
            // -------------------------------------------------
            sheet.Cells["G1"].PutValue("Month");
            sheet.Cells["H1"].PutValue("Sales");
            sheet.Cells["G2"].PutValue("Jan");
            sheet.Cells["H2"].PutValue(150);
            sheet.Cells["G3"].PutValue("Feb");
            sheet.Cells["H3"].PutValue(200);
            sheet.Cells["G4"].PutValue("Mar");
            sheet.Cells["H4"].PutValue(180);
            sheet.Cells["G5"].PutValue("Apr");
            sheet.Cells["H5"].PutValue(220);

            // Add the third chart (Line with Data Markers) and position it on the sheet
            int lineChartIndex = sheet.Charts.Add(ChartType.LineWithDataMarkers, 22, 0, 36, 5);
            Chart lineChart = sheet.Charts[lineChartIndex];
            lineChart.SetChartDataRange("G1:H5", true);
            lineChart.Title.Text = "Line Chart Example";

            // Save the workbook to a file
            workbook.Save("MultipleChartsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
