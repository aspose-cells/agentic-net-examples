// Title: Convert a Range to a ListObject and Bind It to a Column Chart with Aspose.Cells for .NET
// Description: Creates a workbook, fills A1:C5 with headers and data, adds a ListObject (table) over that range, inserts a column chart, and links the chart to the table's DataRange so the visual updates automatically when the table changes.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table | dynamic chart source | column chart | automatic chart update | SetChartDataRange | Excel automation
// Common Searches: Aspose.Cells convert range to ListObject | bind Excel table to chart Aspose.Cells | dynamic chart data source .NET | C# Aspose.Cells chart from table | auto‑updating chart Aspose.Cells
// Developer Intent: Generate a ListObject from a worksheet range and connect it to a column chart so the chart reflects any future modifications to the table data.
// Use Cases: Create a sales table and display live column chart updates as new figures are added. | Reuse a single ListObject as the data source for multiple chart types without redefining cell addresses. | Export a reporting workbook that contains a table and a linked chart for downstream analysis.
// AI Prompts: Write C# code using Aspose.Cells to turn the range A1:D10 into a ListObject named "ReportTable" and add a line chart that references the table's DataRange. | Show how updating values in a ListObject automatically refreshes a linked column chart in an Aspose.Cells workbook. | Provide an example that adds several charts (column, pie, line) all using the same ListObject as their data source in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills A1:C5 with headers and data, adds a ListObject (table) over that range, inserts a column chart, and links the chart to the table's DataRange so the visual updates automatically when the table changes.
    class ConvertRangeToListObjectAndChart
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data (including headers)
            // A1:C1 – headers, A2:C5 – data
            cells["A1"].PutValue("Category");
            cells["B1"].PutValue("Series1");
            cells["C1"].PutValue("Series2");

            for (int i = 2; i <= 5; i++)
            {
                cells[$"A{i}"].PutValue("Cat" + (i - 1));
                cells[$"B{i}"].PutValue(i * 10);
                cells[$"C{i}"].PutValue(i * 15);
            }

            // Define the range that will become a ListObject (table)
            int startRow = cells["A1"].Row;          // 0
            int startColumn = cells["A1"].Column;    // 0
            int endRow = cells["C5"].Row;            // 4
            int endColumn = cells["C5"].Column;      // 2

            // Add the ListObject (table) to the worksheet
            int tableIndex = sheet.ListObjects.Add(startRow, startColumn, endRow, endColumn, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SampleTable";

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart chart = sheet.Charts[chartIndex];

            // Use the table's data range for the chart.
            // DataRange includes the header row; SetChartDataRange expects a string address.
            string dataArea = table.DataRange.Address; // e.g., "A1:C5"
            chart.SetChartDataRange(dataArea, true);   // true = series are plotted by column

            // Optional: set chart title
            chart.Title.Text = "Sample Chart from ListObject";

            // Save the workbook
            workbook.Save("RangeToListObjectChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
