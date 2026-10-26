// Title: Create a dynamic column chart in Aspose.Cells by binding its series to a ListObject built from an in‑memory C# List
// Common Searches: Aspose.Cells bind chart series to Excel table created from List<T> in C# | How to make a column chart update automatically when underlying List data changes using Aspose.Cells | Set chart data source to a ListObject in Aspose.Cells .NET | Create dynamic Excel chart from in‑memory collection with Aspose.Cells | Use Aspose.Cells ListObject as data source for column chart in C#
// Tags: Aspose.Cells chart series binding to ListObject | dynamic Excel column chart from C# collection | programmatic creation of Excel table Aspose.Cells | chart data range address using ListObject Aspose.Cells | in‑memory list as chart source Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;   // For Chart, ChartType
using Aspose.Cells.Tables;   // For ListObject

// The example creates a Workbook, writes a List<SalesData> to cells, converts the range into a ListObject named "SalesTable", adds a column chart, binds the chart's series to the table's data range via its address, and saves the workbook as DynamicChart.xlsx, enabling the chart to reflect changes in the in‑memory list.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // In‑memory list of objects that will drive the chart
            var data = new List<SalesData>
            {
                new SalesData { Month = "Jan", Sales = 120 },
                new SalesData { Month = "Feb", Sales = 150 },
                new SalesData { Month = "Mar", Sales = 180 }
            };

            // Write the list to the worksheet starting at A1
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            for (int i = 0; i < data.Count; i++)
            {
                sheet.Cells[i + 1, 0].PutValue(data[i].Month);   // Column A
                sheet.Cells[i + 1, 1].PutValue(data[i].Sales);   // Column B
            }

            // Convert the range into a ListObject (Excel table) – this makes the data source dynamic
            int totalRows = data.Count + 1;   // header + data rows
            int totalColumns = 2;            // Month and Sales
            int tableIndex = sheet.ListObjects.Add(0, 0, totalRows, totalColumns, true);
            ListObject table = sheet.ListObjects[tableIndex];
            table.DisplayName = "SalesTable";

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Bind the chart series to the table columns using the range address string
            chart.NSeries.Add(table.DataRange.Address, true);

            // Save the workbook (lifecycle rule)
            workbook.Save("DynamicChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Simple POCO representing a row of data
    public class SalesData
    {
        public string Month { get; set; } = string.Empty;
        public double Sales { get; set; }
    }
}
