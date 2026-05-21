using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;   // Required for Chart, ChartType, ChartCollection

namespace AsposeCellsChartProcessing
{
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a sample workbook with a chart (optional, just for demo)
                // ------------------------------------------------------------
                Workbook sourceWorkbook = new Workbook();
                Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

                // Populate some data for the chart
                sourceSheet.Cells["A1"].PutValue("Category");
                sourceSheet.Cells["A2"].PutValue("A");
                sourceSheet.Cells["A3"].PutValue("B");
                sourceSheet.Cells["A4"].PutValue("C");
                sourceSheet.Cells["B1"].PutValue("Value");
                sourceSheet.Cells["B2"].PutValue(10);
                sourceSheet.Cells["B3"].PutValue(20);
                sourceSheet.Cells["B4"].PutValue(30);

                // Add a chart to the worksheet
                int chartIdx = sourceSheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sourceSheet.Charts[chartIdx];
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // ------------------------------------------------------------
                // 2. Save the workbook to a memory stream using the provided method
                // ------------------------------------------------------------
                using (MemoryStream memoryStream = sourceWorkbook.SaveToStream())
                {
                    // Reset the stream position before reading
                    memoryStream.Position = 0;

                    // ------------------------------------------------------------
                    // 3. Load a new workbook from the memory stream using the Stream constructor
                    // ------------------------------------------------------------
                    Workbook loadedWorkbook = new Workbook(memoryStream);

                    // ------------------------------------------------------------
                    // 4. Access the chart collection of the first worksheet for processing
                    // ------------------------------------------------------------
                    Worksheet firstSheet = loadedWorkbook.Worksheets[0];
                    ChartCollection charts = firstSheet.Charts;

                    // Example processing: list chart types and their indexes
                    Console.WriteLine($"Number of charts in the first worksheet: {charts.Count}");
                    for (int i = 0; i < charts.Count; i++)
                    {
                        Chart c = charts[i];
                        Console.WriteLine($"Chart #{i} - Type: {c.Type}");
                        // Additional processing can be done here, e.g., modify series, titles, etc.
                    }

                    // ------------------------------------------------------------
                    // 5. (Optional) Save the processed workbook to verify changes
                    // ------------------------------------------------------------
                    string outputPath = "ProcessedWorkbook.xlsx";
                    loadedWorkbook.Save(outputPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Processed workbook saved to '{Path.GetFullPath(outputPath)}'.");
                }

                Console.WriteLine("Processing completed.");
            }
            catch (FileNotFoundException fnfEx)
            {
                Console.Error.WriteLine($"File not found: {fnfEx.FileName}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}