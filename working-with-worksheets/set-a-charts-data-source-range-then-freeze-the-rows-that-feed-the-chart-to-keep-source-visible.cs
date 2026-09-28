// Title: Bind a column chart to a cell range and freeze the source rows in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Create a column chart that uses cells A2:B11 as its data source, then apply FreezePanes to lock rows 1‑11 so the chart source stays visible, using Aspose.Cells in C#. | Generate sample data in columns A and B, add a column chart referencing that data, freeze the first 11 rows of the worksheet, and save the workbook as an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells C# set chart data source range and freeze rows | freeze panes for chart source data using Aspose.Cells .NET | C# Aspose.Cells column chart from A2:B11 with frozen source rows | keep chart source rows visible in Excel with Aspose.Cells
// Tags: Aspose.Cells set chart data source range | Aspose.Cells FreezePanes rows example | Aspose.Cells column chart from worksheet range | Aspose.Cells save workbook with frozen rows | Aspose.Cells populate worksheet sample data

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a new workbook, fills cells A1:B11 with sample categories and values, adds a column chart bound to the range A2:B11, freezes the first 11 rows to keep the chart's source data visible, and saves the file as ChartWithFrozenRows.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data that will be used as the chart source
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 11; i++) // rows 2‑11 contain data
            {
                sheet.Cells[i - 1, 0].PutValue("Item " + (i - 1)); // Column A
                sheet.Cells[i - 1, 1].PutValue(i * 10);           // Column B
            }

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 13, 0, 30, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source range
            // Values series: B2:B11, Categories: A2:A11
            chart.NSeries.Add("=Sheet1!$B$2:$B$11", true);
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$11";

            // Freeze the rows that contain the source data (rows 1‑11)
            // FreezePanes(row, column, totalRows, totalColumns) uses zero‑based indexes.
            // To freeze the first 11 rows, set row = 11 and column = 0.
            sheet.FreezePanes(11, 0, 0, 0);

            // Define output file path
            string outputPath = "ChartWithFrozenRows.xlsx";

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
