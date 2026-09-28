// Title: Generate a column chart from a worksheet ListObject table and apply a built‑in chart style using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that creates a ListObject table, assigns a DisplayName, and builds a column chart whose series reference the table columns via the DisplayName syntax. | Demonstrate how to assign one of Aspose.Cells' predefined chart styles to a column chart after its series are linked to a worksheet table.
// Common Searches: how to use a ListObject as data source for a chart in Aspose.Cells C# | Aspose.Cells bind column chart series to table columns example | set a predefined chart style on an Aspose.Cells chart using C# | C# create worksheet table and chart from it using Aspose.Cells | Aspose.Cells chart NSeries reference table column syntax
// Tags: Aspose.Cells create ListObject table | Aspose.Cells column chart from table | Aspose.Cells set chart style programmatically | Aspose.Cells NSeries table column reference | Aspose.Cells C# chart data binding

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

// The program creates a new workbook, fills it with sample data, defines a ListObject table named DataTable, adds a column chart whose series are linked to the table columns using the DisplayName syntax, applies a predefined chart style, and saves the file as ChartFromTable.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value1");
            sheet.Cells["C1"].PutValue("Value2");

            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["C2"].PutValue(20);

            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["C3"].PutValue(40);

            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(50);
            sheet.Cells["C4"].PutValue(60);

            // Create a table (ListObject) that covers the data range
            // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
            int tableIndex = sheet.ListObjects.Add(0, 0, 4, 3, true);
            ListObject table = sheet.ListObjects[tableIndex];
            // Use DisplayName to set the table name (Name property is not available in this version)
            table.DisplayName = "DataTable";

            // Add a column chart to the worksheet
            // Parameters: chart type, upper left row, upper left column, lower right row, lower right column
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source to the table columns
            // Category (X) axis from the "Category" column of the table
            chart.NSeries.CategoryData = $"{table.DisplayName}[Category]";

            // Add two series using the "Value1" and "Value2" columns of the table
            chart.NSeries.Add($"{table.DisplayName}[Value1]", true);
            chart.NSeries.Add($"{table.DisplayName}[Value2]", true);

            // Save the workbook
            workbook.Save("ChartFromTable.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
