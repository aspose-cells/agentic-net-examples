// Title: Export a chart series range to a JSON array using Aspose.Cells JsonUtility in C#
// AI Prompts: Write C# code that obtains the data range of a chart series in an Aspose.Cells workbook and calls JsonUtility.ExportRangeToJson with appropriate JsonSaveOptions to produce a JSON array. | Show how to include both the category (X) and value (Y) fields in the exported JSON by exporting the chart's category range together with its value range. | Provide a C# example that writes the JSON string returned by JsonUtility.ExportRangeToJson to a .json file while leaving the original Excel workbook untouched.
// Common Searches: how to use Aspose.Cells JsonUtility to export chart series data as JSON in C# | C# code example for exporting Excel chart values to a JSON array with Aspose.Cells | Aspose.Cells JsonSaveOptions settings for exporting a range without a header row | extract X and Y values from a chart series and save them to JSON using Aspose.Cells | export chart series range to JSON file while keeping workbook unchanged Aspose.Cells
// Tags: Aspose.Cells chart series JSON export | JsonUtility ExportRangeToJson C# | JsonSaveOptions without header row | retrieve chart series range Aspose.Cells | convert Excel chart data to JSON array

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Utility;

// Alias to avoid conflict with System.Range
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, fills columns with category and value data, adds a column chart bound to those cells, defines the series range, configures JsonSaveOptions to export values as strings without a header, and uses JsonUtility.ExportRangeToJson to generate a JSON array from the range. The JSON is printed to the console, and the workbook is optionally saved.
class ExportChartSeriesToJson
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Fill sample data: categories in column A, values in column B
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["B1"].PutValue("Value");
            for (int i = 2; i <= 6; i++)
            {
                worksheet.Cells[$"A{i}"].PutValue("Cat " + (i - 1));
                worksheet.Cells[$"B{i}"].PutValue(i * 10);
            }

            // Add a column chart and bind its series to the values column (B2:B6)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 5);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("=Sheet1!$B$2:$B$6", true);               // Y‑values
            chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$6";          // X‑axis categories

            // Define the range that holds the series data (column B, rows 2‑6)
            // Row and column indices are zero‑based: startRow=1, startColumn=1, rowCount=5, columnCount=1
            AsposeRange seriesRange = worksheet.Cells.CreateRange(1, 1, 5, 1);

            // Set JSON export options (export values as strings, no header row)
            JsonSaveOptions jsonOptions = new JsonSaveOptions
            {
                ExportAsString = true,
                HasHeaderRow = false,
                ExportEmptyCells = false
            };

            // Export the range to a JSON string
            string jsonResult = JsonUtility.ExportRangeToJson(seriesRange, jsonOptions);

            // Output the JSON array
            Console.WriteLine(jsonResult);

            // Save the workbook (optional, not required for JSON export)
            string outputPath = "ChartSeriesDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
