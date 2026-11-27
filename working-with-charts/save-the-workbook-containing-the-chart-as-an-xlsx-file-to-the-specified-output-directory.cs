// Title: Create a column chart in a new workbook and save it as an XLSX file to a custom output folder using Aspose.Cells for .NET
// Common Searches: how to generate a column chart with Aspose.Cells and write the workbook to a specific directory in C# | Aspose.Cells C# save workbook containing a chart to a user-defined folder | C# example for creating a chart and exporting the workbook as XLSX with Aspose.Cells | saving an Aspose.Cells workbook with an embedded chart to a dynamic output path | Aspose.Cells .NET create chart and specify SaveFormat.Xlsx location
// Tags: Aspose.Cells create column chart C# | Aspose.Cells save workbook to XLSX | Aspose.Cells specify output directory | Aspose.Cells chart NSeries binding | Aspose.Cells workbook.Save with SaveFormat.Xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSaveExample
{
    // The sample determines an output folder (default "Output" or a path passed via command line), creates the directory if needed, builds a new workbook, fills cells A1:A5 and B1:B5 with category labels and numeric values, adds a column chart linked to that data, sets a title and legend position, and finally saves the workbook as "ChartWorkbook.xlsx" in XLSX format to the chosen directory using Aspose.Cells for .NET.
    class Program
    {
        static int Main(string[] args)
        {
            // Determine the output directory (first argument or default "Output")
            string outputDirectory = args.Length > 0 ? args[0] : "Output";

            try
            {
                // Ensure the output directory exists
                Directory.CreateDirectory(outputDirectory);

                // Define the full path for the resulting XLSX file
                string outputPath = Path.Combine(outputDirectory, "ChartWorkbook.xlsx");

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                // A1:A5 - Categories, B1:B5 - Values
                worksheet.Cells["A1"].PutValue("Category 1");
                worksheet.Cells["A2"].PutValue("Category 2");
                worksheet.Cells["A3"].PutValue("Category 3");
                worksheet.Cells["A4"].PutValue("Category 4");
                worksheet.Cells["A5"].PutValue("Category 5");

                worksheet.Cells["B1"].PutValue(10);
                worksheet.Cells["B2"].PutValue(20);
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["B4"].PutValue(25);
                worksheet.Cells["B5"].PutValue(15);

                // Add a column chart to the worksheet
                // Parameters: ChartType, upper left row, upper left column, lower right row, lower right column
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 7, 0, 20, 7);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data series (values) and category (X axis) data
                chart.NSeries.Add("B1:B5", true);
                chart.NSeries.CategoryData = "A1:A5";

                // Set chart title
                chart.Title.Text = "Sample Column Chart";

                // Optional: format the chart (e.g., set legend position)
                chart.Legend.Position = LegendPositionType.Right;

                // Save the workbook as XLSX
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine($"Workbook with chart saved successfully to: {outputPath}");
                return 0;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine("An error occurred while creating or saving the workbook:");
                Console.Error.WriteLine(ex.Message);
                return 1;
            }
        }
    }
}
