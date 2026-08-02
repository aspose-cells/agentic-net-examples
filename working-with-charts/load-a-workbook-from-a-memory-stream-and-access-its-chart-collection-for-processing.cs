using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartProcessing
{
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a sample workbook and add a chart
            // -------------------------------------------------
            Workbook sourceWorkbook = new Workbook();

            // Populate some data for the chart
            Worksheet srcSheet = sourceWorkbook.Worksheets[0];
            srcSheet.Cells["A1"].PutValue("Category");
            srcSheet.Cells["A2"].PutValue("A");
            srcSheet.Cells["A3"].PutValue("B");
            srcSheet.Cells["A4"].PutValue("C");
            srcSheet.Cells["B1"].PutValue("Value");
            srcSheet.Cells["B2"].PutValue(10);
            srcSheet.Cells["B3"].PutValue(20);
            srcSheet.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIdx = srcSheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = srcSheet.Charts[chartIdx];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // -------------------------------------------------
            // 2. Save the workbook to a memory stream (xls format)
            // -------------------------------------------------
            MemoryStream stream = sourceWorkbook.SaveToStream();

            // Reset stream position before reading
            stream.Position = 0;

            // -------------------------------------------------
            // 3. Load a workbook from the memory stream
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(stream);

            // -------------------------------------------------
            // 4. Access the chart collection of the first worksheet
            // -------------------------------------------------
            Worksheet firstSheet = loadedWorkbook.Worksheets[0];
            ChartCollection charts = firstSheet.Charts;

            // Example processing: list chart types and their indexes
            Console.WriteLine($"Number of charts in the worksheet: {charts.Count}");
            for (int i = 0; i < charts.Count; i++)
            {
                Chart c = charts[i];
                Console.WriteLine($"Chart #{i} - Type: {c.Type}");
                // Additional processing can be added here (e.g., modify series, titles, etc.)
            }

            // Clean up
            stream.Dispose();
            sourceWorkbook.Dispose();
            loadedWorkbook.Dispose();
        }
    }
}