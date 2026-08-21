// Title: C# – Create a Column Chart from a Worksheet Table and Apply a Built‑In Chart Style with Aspose.Cells
// Description: This Aspose.Cells for .NET example shows how to build a new workbook, turn a range into a worksheet table (ListObject), bind the table columns to a column chart, apply a built‑in chart style (e.g., style 2), set a title, and save the file as an XLSX workbook.
// Keywords: Aspose.Cells C# chart example | create chart from ListObject | worksheet table chart source .NET | built‑in chart style Aspose.Cells | column chart styling Excel automation | Aspose.Cells chart style number | Excel table to chart Aspose | Aspose.Cells sample code | C# Excel chart generation | Aspose.Cells chart template
// Common Searches: Aspose.Cells bind worksheet table to chart | apply built‑in chart style in Aspose.Cells C# | create column chart from ListObject Aspose | Aspose.Cells chart style numbers | C# example chart from Excel table Aspose.Cells | how to use ListObject as chart source Aspose
// Developer Intent: Generate a column chart that reads data from a worksheet table and apply a predefined built‑in chart style.
// Use Cases: Automated sales dashboards that pull values from a table and use a corporate chart style. | Reusable Excel report templates where chart data updates automatically as the underlying ListObject changes. | Batch creation of styled charts for financial summaries across multiple workbooks.
// AI Prompts: Show how to change the chart to a different built‑in style number and set the title from a variable. | Add data labels, a legend, and a secondary axis to the chart that uses the ListObject as its source. | Explain how to reference a table column for a line chart’s NSeries in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This Aspose.Cells for .NET example shows how to build a new workbook, turn a range into a worksheet table (ListObject), bind the table columns to a column chart, apply a built‑in chart style (e.g., style 2), set a title, and save the file as an XLSX workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(30);

        // Convert the range A1:B4 into a worksheet table (ListObject)
        // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
        int tableIdx = sheet.ListObjects.Add(0, 0, 4, 2, true);
        // Assign a friendly name to the table for easier reference
        sheet.ListObjects[tableIdx].DisplayName = "SalesTable";

        // Add a column chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
        Chart chart = sheet.Charts[chartIdx];

        // Use the table columns as the chart data source
        chart.NSeries.Add("SalesTable[Value]", true);          // Values series
        chart.NSeries.CategoryData = "SalesTable[Category]"; // Category (X‑axis) data

        // Apply a predefined built‑in chart style (style numbers range from 1 to 48)
        chart.Style = 2; // Example: style 2

        // Optional: set a chart title
        chart.Title.Text = "Sales Chart";

        // Save the workbook with the chart
        workbook.Save("ChartWithTableStyle.xlsx");
    }
}
