// Title: Export Aspose.Cells Chart Series to JSON in C# by Iterating Series Ranges
// Description: Creates a workbook, adds a column chart, loops through each NSeries, parses the series.Values formula, uses JsonUtility.ExportRangeToJson with JsonSaveOptions (ExportAsString=true, HasHeaderRow=false) and combines the results into a single JSON array.
// Keywords: Aspose.Cells | C# chart to JSON | ExportRangeToJson | JsonUtility | JsonSaveOptions | chart series extraction | Aspose.Cells NSeries iteration
// Common Searches: Aspose.Cells export chart series to JSON C# | How to convert chart data to JSON with Aspose.Cells | Iterate chart NSeries and export ranges as JSON | JsonUtility ExportRangeToJson example for charts | C# Aspose.Cells chart data JSON output
// Developer Intent: Extract the data points of each chart series and output them as a JSON array.
// Use Cases: Provide a JSON endpoint for JavaScript chart libraries to replicate an Excel chart. | Store chart series values in JSON for analytics or archival purposes. | Integrate Excel chart data into web APIs or reporting dashboards.
// AI Prompts: Generate C# code that iterates over an Aspose.Cells chart's NSeries, extracts the source range from each series.Values formula, and exports the range to JSON using JsonUtility with ExportAsString and no header row. | Show how to merge multiple JSON strings representing individual chart series into a single JSON array in C#. | Explain the configuration of JsonSaveOptions for exporting cell values as strings and omitting header rows when using Aspose.Cells JsonUtility.

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartJsonExport
{
    // Creates a workbook, adds a column chart, loops through each NSeries, parses the series.Values formula, uses JsonUtility.ExportRangeToJson with JsonSaveOptions (ExportAsString=true, HasHeaderRow=false) and combines the results into a single JSON array.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, two series in columns B and C)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["C1"].PutValue("Series2");

                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a column chart and bind the two series
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("=Sheet1!$B$2:$B$4", true); // Series1 values
                chart.NSeries.Add("=Sheet1!$C$2:$C$4", true); // Series2 values
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";

                // Prepare JSON export options (export as string values, exclude header row)
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    ExportAsString = true,
                    HasHeaderRow = false
                };

                // Collect JSON for each series
                List<string> seriesJsonList = new List<string>();
                foreach (Series series in chart.NSeries)
                {
                    // Extract the range address from the series.Values formula (e.g., "=Sheet1!$B$2:$B$4")
                    string formula = series.Values;
                    if (string.IsNullOrEmpty(formula))
                        continue;

                    // Remove leading '=' and sheet name
                    int exclPos = formula.IndexOf('!');
                    string address = exclPos >= 0 ? formula.Substring(exclPos + 1) : formula;
                    // Remove any '$' characters
                    address = address.Replace("$", string.Empty);

                    // Create the range object from the address
                    AsposeRange dataRange = sheet.Cells.CreateRange(address);

                    // Export the range to JSON string
                    string json = JsonUtility.ExportRangeToJson(dataRange, jsonOptions);
                    seriesJsonList.Add(json);
                }

                // Combine individual series JSON strings into a single JSON array
                string finalJson = "[" + string.Join(",", seriesJsonList) + "]";

                // Output the combined JSON
                Console.WriteLine("Chart series data exported to JSON:");
                Console.WriteLine(finalJson);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
