// Title: Create Multiple Charts (Column, Pie, Line) on a Single Worksheet with Aspose.Cells for .NET (C#)
// Description: C# sample that builds a new workbook, populates three separate data tables, and adds a column chart, a pie chart, and a line chart—each bound to its own range and positioned independently on the same sheet. The workbook is saved as an XLSX file using the Aspose.Cells API.
// Keywords: Aspose.Cells | C# | .NET | multiple charts | column chart | pie chart | line chart | SetChartDataRange | chart positioning | Excel dashboard example | sample code | GitHub Aspose.Cells | Excel automation
// Common Searches: Aspose.Cells add several charts to one worksheet C# | set data range for multiple charts Aspose.Cells .NET | create column, pie and line charts on same sheet using Aspose.Cells | position multiple charts in Excel with Aspose.Cells | sample code for multiple chart types Aspose.Cells C#
// Developer Intent: Add three distinct charts—column, pie, and line—to a single worksheet, each with its own data range and layout.
// Use Cases: Generate a sales dashboard that shows monthly sales (column), product mix (pie), and weekly temperature trends (line) on one sheet. | Automate Excel reports that combine financial, inventory, and operational visualizations without manual chart placement. | Create a reusable template for multi‑chart worksheets in enterprise reporting solutions.
// AI Prompts: Show how to add an area chart as a fourth visual and bind it to a new data range. | Provide code to calculate chart positions dynamically so any number of charts can be added without overlap. | Explain how to customize colors, markers, and legends for each chart created with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsMultipleChartsDemo
{
    // C# sample that builds a new workbook, populates three separate data tables, and adds a column chart, a pie chart, and a line chart—each bound to its own range and positioned independently on the same sheet. The workbook is saved as an XLSX file using the Aspose.Cells API.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "DataSheet";

            // -------------------------------------------------
            // Prepare sample data for three different charts
            // -------------------------------------------------

            // Data for Chart 1 (Column Chart)
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            string[] months = { "Jan", "Feb", "Mar", "Apr", "May" };
            int[] sales = { 1200, 1500, 1800, 1300, 1700 };
            for (int i = 0; i < months.Length; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(months[i]);   // Column A
                sheet.Cells[i + 1, 1].PutValue(sales[i]);   // Column B
            }

            // Data for Chart 2 (Pie Chart)
            sheet.Cells["D1"].PutValue("Product");
            sheet.Cells["E1"].PutValue("Quantity");
            string[] products = { "A", "B", "C", "D" };
            int[] qty = { 40, 30, 20, 10 };
            for (int i = 0; i < products.Length; i++)
            {
                sheet.Cells[i + 1, 3].PutValue(products[i]); // Column D
                sheet.Cells[i + 1, 4].PutValue(qty[i]);     // Column E
            }

            // Data for Chart 3 (Line Chart)
            sheet.Cells["G1"].PutValue("Day");
            sheet.Cells["H1"].PutValue("Temperature");
            for (int i = 1; i <= 7; i++)
            {
                sheet.Cells[i, 6].PutValue("Day " + i);          // Column G
                sheet.Cells[i, 7].PutValue(15 + i * 2);          // Column H
            }

            // -------------------------------------------------
            // Add Chart 1: Column Chart using range A1:B6
            // -------------------------------------------------
            int chartIndex1 = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 7);
            Chart chart1 = sheet.Charts[chartIndex1];
            chart1.Title.Text = "Monthly Sales";
            chart1.SetChartDataRange("A1:B6", true); // Plot by column

            // -------------------------------------------------
            // Add Chart 2: Pie Chart using range D1:E5
            // -------------------------------------------------
            int chartIndex2 = sheet.Charts.Add(ChartType.Pie, 10, 9, 25, 16);
            Chart chart2 = sheet.Charts[chartIndex2];
            chart2.Title.Text = "Product Distribution";
            chart2.SetChartDataRange("D1:E5", true);

            // -------------------------------------------------
            // Add Chart 3: Line Chart using range G1:H8
            // -------------------------------------------------
            int chartIndex3 = sheet.Charts.Add(ChartType.Line, 30, 0, 45, 7);
            Chart chart3 = sheet.Charts[chartIndex3];
            chart3.Title.Text = "Weekly Temperature";
            chart3.SetChartDataRange("G1:H8", true);

            // -------------------------------------------------
            // Save the workbook
            // -------------------------------------------------
            workbook.Save("MultipleChartsDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
