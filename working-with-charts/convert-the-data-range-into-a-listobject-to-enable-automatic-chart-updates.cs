// Title: Create a ListObject from a worksheet range and bind it to a column chart for auto‑updating data in Aspose.Cells C#
// AI Prompts: Generate C# code that converts cells A1:C5 into a ListObject named SalesTable and links the table’s DataRange to a column chart using Aspose.Cells. | Write a C# snippet that adds a ListObject to a worksheet, sets its display name, creates a column chart, and assigns the chart’s data source to the table so the chart refreshes when the table data changes.
// Common Searches: aspnet how to create a ListObject from a range and use it as chart source in Aspose.Cells | c# Aspose.Cells column chart automatically updates when ListObject data changes | set chart data range to ListObject DataRange Aspose.Cells example | convert worksheet range to ListObject for dynamic chart in Aspose.Cells .NET
// Tags: Aspose.Cells ListObject creation C# | Aspose.Cells bind chart to ListObject | Aspose.Cells dynamic chart data source | Aspose.Cells column chart from table | Aspose.Cells auto‑refresh chart with table changes

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsListObjectChartDemo
{
    // The program creates a workbook, fills cells A1:C5 with sample data, converts that range into a ListObject named "SalesTable", adds a column chart, sets the chart's data source to the table's DataRange, and saves the file as ListObjectChartDemo.xlsx.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data with headers (A1:C5)
            // Header row
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Series1");
            worksheet.Cells["C1"].PutValue("Series2");

            // Data rows
            worksheet.Cells["A2"].PutValue("Jan");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["C2"].PutValue(15);

            worksheet.Cells["A3"].PutValue("Feb");
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["C3"].PutValue(25);

            worksheet.Cells["A4"].PutValue("Mar");
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["C4"].PutValue(35);

            worksheet.Cells["A5"].PutValue("Apr");
            worksheet.Cells["B5"].PutValue(40);
            worksheet.Cells["C5"].PutValue(45);

            // Convert the data range into a ListObject (table) so that charts update automatically
            // Using the Add(string, string, bool) overload as defined in the rules
            int tableIndex = worksheet.ListObjects.Add("A1", "C5", true);
            ListObject table = worksheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the chart's data range to the table's data range.
            // The DataRange.Address returns the address string (e.g., "A1:C5").
            // Using SetChartDataRange(string area, bool isVertical) as defined in the rules.
            chart.SetChartDataRange(table.DataRange.Address, true);

            // Optional: give the chart a title
            chart.Title.Text = "Monthly Sales";

            // Save the workbook
            workbook.Save("ListObjectChartDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}
