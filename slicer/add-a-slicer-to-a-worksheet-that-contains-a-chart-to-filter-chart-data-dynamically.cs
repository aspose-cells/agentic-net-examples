// Title: Create an Excel table and linked column chart with Aspose.Cells for .NET (C#)
// Common Searches: how to generate an Excel table and column chart using Aspose.Cells in C# | Aspose.Cells example for adding a ListObject and chart programmatically | C# code to create a column chart from a table range with Aspose.Cells | using Aspose.Cells .NET to build a workbook with data table and chart | sample Aspose.Cells project that creates a chart from a ListObject
// Tags: Aspose.Cells create table and chart | C# programmatic Excel ListObject with chart | Aspose.Cells column chart from table range | Excel workbook generation using Aspose.Cells .NET | dynamic data visualization with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;
using Aspose.Cells.Charts;

namespace AsposeCellsSlicerExample
{
    // The example creates a new Excel workbook, adds a ListObject (table) with sample category and value data, inserts a column chart that references the table range for series and categories, and saves the file as SlicerChartExample.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet and give it a friendly name
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "DataSheet";

                // Populate sample data (Category and Value)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["A5"].PutValue("D");
                sheet.Cells["B5"].PutValue(40);
                sheet.Cells["A6"].PutValue("E");
                sheet.Cells["B6"].PutValue(50);

                // Define the data range for the table (including headers)
                int firstRow = 0;          // zero‑based index for row 1
                int firstColumn = 0;       // column A
                int totalRows = 6;         // rows 1‑6
                int totalColumns = 2;      // columns A‑B

                // Add a ListObject (table) to the worksheet
                int tableIndex = sheet.ListObjects.Add(firstRow, firstColumn,
                    firstRow + totalRows - 1, firstColumn + totalColumns - 1, true);
                ListObject table = sheet.ListObjects[tableIndex];
                table.DisplayName = "SalesTable";
                table.ShowHeaderRow = true;
                table.ShowTableStyleColumnStripes = true;
                table.TableStyleType = TableStyleType.TableStyleMedium9;

                // Add a column chart that uses the table data
                // Height and width are specified in points (rows, columns)
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 15, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.Title.Text = "Sales by Category";

                // Use the table range for the chart series and categories
                // This avoids reliance on the ListObject.Name property
                string valueRange = $"'{sheet.Name}'!B2:B6";
                string categoryRange = $"'{sheet.Name}'!A2:A6";
                chart.NSeries.Add(valueRange, true);
                chart.NSeries.CategoryData = categoryRange;

                // Save the workbook to a file (ensure the directory exists)
                string outputPath = "SlicerChartExample.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
                }
                catch (Exception saveEx)
                {
                    Console.WriteLine($"Error saving workbook: {saveEx.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
