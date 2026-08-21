// Title: Convert a Range to a ListObject (Table) and Bind a Column Chart for Auto‑Update – Aspose.Cells for .NET
// Description: C# example that creates a workbook, fills A1:C5 with headers and data, transforms the range into a ListObject named "SalesTable", adds a column chart, and sets the chart's source to the table's DataRange so the visual refreshes automatically when the table is edited. The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | ListObject | Excel table | column chart | automatic chart refresh | SetChartDataRange | dynamic chart source | Excel automation
// Common Searches: Aspose.Cells bind chart to ListObject | convert cell range to table Aspose.Cells C# | auto‑update chart when table changes Aspose.Cells | SetChartDataRange using ListObject.DataRange | create column chart from Excel table Aspose
// Developer Intent: Create a ListObject from a worksheet range and link a column chart to the table so the chart updates automatically as the table data changes.
// Use Cases: Sales reporting workbook where new rows added to the table instantly reflect in the column chart. | Interactive dashboard that uses multiple tables as data sources, keeping all charts synchronized with edits. | Exporting a financial model where filtered or sorted table data drives charts without manual range adjustments.
// AI Prompts: Generate C# code that adds a ListObject to a worksheet and binds a column chart to its DataRange using Aspose.Cells. | Show how to make an Aspose.Cells chart refresh automatically when rows are added to a ListObject. | Explain the steps to set a chart's data source to a table address and ensure dynamic updates in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

// C# example that creates a workbook, fills A1:C5 with headers and data, transforms the range into a ListObject named "SalesTable", adds a column chart, and sets the chart's source to the table's DataRange so the visual refreshes automatically when the table is edited. The workbook is saved as an XLSX file.
class ListObjectChartDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data with headers (A1:C5)
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["C1"].PutValue("Series2");

        sheet.Cells["A2"].PutValue("Jan");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["C2"].PutValue(15);

        sheet.Cells["A3"].PutValue("Feb");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["C3"].PutValue(25);

        sheet.Cells["A4"].PutValue("Mar");
        sheet.Cells["B4"].PutValue(30);
        sheet.Cells["C4"].PutValue(35);

        sheet.Cells["A5"].PutValue("Apr");
        sheet.Cells["B5"].PutValue(40);
        sheet.Cells["C5"].PutValue(45);

        // Convert the range into a ListObject (table) so that chart data updates automatically
        int tableIndex = sheet.ListObjects.Add("A1", "C5", true);
        ListObject table = sheet.ListObjects[tableIndex];
        table.DisplayName = "SalesTable"; // optional: give the table a name

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 5);
        Chart chart = sheet.Charts[chartIndex];

        // Set the chart data range to the table's data range.
        // Using the table's address ensures the chart reflects any changes to the table.
        chart.SetChartDataRange(table.DataRange.Address, true);

        // Save the workbook
        workbook.Save("ListObjectChartDemo.xlsx", SaveFormat.Xlsx);
    }
}
