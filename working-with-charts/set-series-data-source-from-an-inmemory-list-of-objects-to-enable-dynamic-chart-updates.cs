// Title: Generate a dynamic Excel column chart in C# with Aspose.Cells by binding a List<DataItem> as the series source
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, uses WorkbookDesigner to bind a List<T> to placeholder markers, expands the data, and automatically sets the chart series range. | Show how to determine the last populated row after data expansion and assign the corresponding cell addresses to NSeries.ValueData and NSeries.CategoryData for a column chart. | Demonstrate refreshing an Aspose.Cells chart when the underlying in‑memory collection changes, including re‑processing the designer and updating the series ranges before saving.
// Common Searches: how to bind a C# List to an Aspose.Cells chart series using placeholders | Aspose.Cells dynamic chart data range from in‑memory collection .NET | C# create Excel column chart with data from POCO list using Aspose.Cells | update Aspose.Cells chart after processing data source programmatically
// Tags: bind collection to Aspose.Cells chart via designer | expand placeholder markers into worksheet rows | assign chart series ranges programmatically | generate column chart from in‑memory data | save dynamic chart as .xlsx using C#

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDynamicChart
{
    // Simple POCO representing a data point
    // The example builds a List<DataItem> as an in‑memory data source, employs Aspose.Cells WorkbookDesigner with placeholder markers (&=$Data.Category, &=$Data.Value) to expand the collection into worksheet rows, adds a column chart, calculates the populated range, sets the NSeries value and category ranges, optionally names the series, and saves the workbook as DynamicChartFromList.xlsx.
    public class DataItem
    {
        public string Category { get; set; }
        public double Value { get; set; }

        public DataItem(string category, double value)
        {
            Category = category;
            Value = value;
        }
    }

    public class Program
    {
        public static void Main()
        {
            try
            {
                // 1. Prepare in‑memory data source
                List<DataItem> data = new List<DataItem>
                {
                    new DataItem("A", 10),
                    new DataItem("B", 20),
                    new DataItem("C", 30),
                    new DataItem("D", 25)
                };

                // 2. Create a new workbook and place smart markers where data will be expanded
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Header row
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");

                // Smart markers – they will be replaced by the designer with the list contents
                sheet.Cells["A2"].PutValue("&=$Data.Category");
                sheet.Cells["B2"].PutValue("&=$Data.Value");

                // 3. Bind the in‑memory list to the smart marker name "Data"
                WorkbookDesigner designer = new WorkbookDesigner
                {
                    Workbook = workbook
                };
                designer.SetDataSource("Data", data);
                designer.Process(); // expands the smart markers into concrete rows

                // 4. Add a chart that uses the populated range
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Determine the last row of data after processing (zero‑based index)
                int lastRow = sheet.Cells.MaxDataRow; // includes header row

                // Build the address strings for values and categories (Excel rows are 1‑based)
                string valuesRange = $"=Sheet1!$B$2:$B${lastRow + 1}";
                string categoryRange = $"=Sheet1!$A$2:$A${lastRow + 1}";

                // Add series using NSeries.Add (string, bool) rule
                chart.NSeries.Add(valuesRange, true);
                // Set category (X‑axis) data
                chart.NSeries.CategoryData = categoryRange;

                // Optional: give the series a name
                chart.NSeries[0].Name = "Sample Series";

                // 5. Save the workbook
                string outputPath = "DynamicChartFromList.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
