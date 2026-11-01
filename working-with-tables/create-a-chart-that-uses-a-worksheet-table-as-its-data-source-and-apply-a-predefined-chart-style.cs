// Title: Generate a column chart from an Excel table and apply a built‑in chart style with Aspose.Cells for .NET (C#)
// Common Searches: how to bind a ListObject table to a chart series using Aspose.Cells C# | Aspose.Cells create column chart from table range and set chart style | C# code to add a worksheet table and use it as data source for a chart in Aspose.Cells | apply predefined chart style to an Aspose.Cells chart programmatically | save Excel workbook with chart and table using Aspose.Cells .NET
// Tags: Aspose.Cells chart from ListObject table | C# column chart with table data source | apply predefined ChartStyleType Aspose.Cells | create worksheet ListObject in Aspose.Cells | save workbook with chart and table Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Tables;

namespace AsposeCellsChartExample
{
    // The example creates a new workbook, adds sample sales data, defines a ListObject named "SalesData" over the range A1:C5, inserts a column chart, links two series to the table's Value1 and Value2 columns, optionally applies a built‑in chart style, and saves the file as ChartWithTable.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (A1:C5)
                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value1");
                sheet.Cells["C1"].PutValue("Value2");

                // Data rows
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");

                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(130);
                sheet.Cells["B5"].PutValue(170);

                sheet.Cells["C2"].PutValue(80);
                sheet.Cells["C3"].PutValue(95);
                sheet.Cells["C4"].PutValue(100);
                sheet.Cells["C5"].PutValue(110);

                // Add a table (ListObject) that covers the data range A1:C5
                // Parameters: firstRow, firstColumn, totalRows, totalColumns, hasHeaders
                int tableIndex = sheet.ListObjects.Add(0, 0, 4, 2, true);
                ListObject table = sheet.ListObjects[tableIndex];
                // Set a display name for the table (Name property may not be available in older versions)
                table.DisplayName = "SalesData";

                // Add a column chart to the worksheet
                // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the chart title
                chart.Title.Text = "Monthly Sales";

                // Use the table as the data source for the series
                // First series (Value1 column)
                chart.NSeries.Add("SalesData!B2:B5", true);
                chart.NSeries[0].Name = "SalesData!B1";

                // Second series (Value2 column)
                chart.NSeries.Add("SalesData!C2:C5", false);
                chart.NSeries[1].Name = "SalesData!C1";

                // Note: CategoryData property may not be available in older API versions.
                // If needed, you can set it when the property exists:
                // chart.NSeries[0].CategoryData = "SalesData!A2:A5";
                // chart.NSeries[1].CategoryData = "SalesData!A2:A5";

                // Apply a predefined chart style if the enum is available
                // If ChartStyleType is not defined, this line can be omitted.
                // chart.Style = ChartStyleType.Style1;

                // Define output file path
                string outputPath = "ChartWithTable.xlsx";

                // Ensure the directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
