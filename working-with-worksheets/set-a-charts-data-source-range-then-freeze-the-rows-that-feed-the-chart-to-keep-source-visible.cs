// Title: How to set a chart’s data source range and freeze the source rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a column chart from cells A2:B4, assigns the series and category ranges, then freezes the first four rows of the worksheet with Aspose.Cells. | Generate a complete Aspose.Cells example that sets NSeries and CategoryData for a chart, applies FreezePanes to keep the chart data rows visible while scrolling, and saves the workbook as an .xlsx file.
// Common Searches: Aspose.Cells C# set chart series range B2:B4 and category range A2:A4 | Freeze first four rows in an Excel worksheet using Aspose.Cells FreezePanes | How to keep chart source data visible while scrolling in Aspose.Cells | C# example for adding a column chart and freezing rows in the same worksheet with Aspose.Cells | Aspose.Cells chart data source and FreezePanes combined example
// Tags: Aspose.Cells chart NSeries range C# | Aspose.Cells FreezePanes rows C# | Aspose.Cells column chart data source | Aspose.Cells keep chart source visible | Aspose.Cells workbook save xlsx C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// The program creates a new workbook, fills sample data, adds a column chart with its series set to B2:B4 and categories to A2:A4, freezes the first four rows of the worksheet to keep the chart source visible, and saves the file as ChartWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used as the chart source
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(15);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source range
            // Values series: B2:B4, Categories: A2:A4
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Freeze the first four rows so they stay visible while scrolling
            // Use the 4‑parameter overload: row, column, totalRows, totalColumns
            int totalRows = sheet.Cells.MaxDataRow + 1;
            int totalColumns = sheet.Cells.MaxDataColumn + 1;
            sheet.FreezePanes(4, 0, totalRows, totalColumns);

            // Save the workbook to a file
            workbook.Save("ChartWithFrozenRows.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
