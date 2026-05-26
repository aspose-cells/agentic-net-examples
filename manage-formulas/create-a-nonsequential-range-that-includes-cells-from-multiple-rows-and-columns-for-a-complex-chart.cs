using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;   // Alias to avoid conflict with System.Range

namespace AsposeCellsComplexChartDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate three separate data blocks (non‑contiguous)
                // Block 1: A1:A5
                for (int i = 0; i < 5; i++)
                    cells[i, 0].PutValue(i + 1);               // Values 1‑5

                // Block 2: C1:C5
                for (int i = 0; i < 5; i++)
                    cells[i, 2].PutValue((i + 1) * 10);        // Values 10‑50

                // Block 3: E1:E5
                for (int i = 0; i < 5; i++)
                    cells[i, 4].PutValue((i + 1) * 100);       // Values 100‑500

                // Create individual Range objects for each block
                AsposeRange range1 = cells.CreateRange("A1", "A5");
                AsposeRange range2 = cells.CreateRange("C1", "C5");
                AsposeRange range3 = cells.CreateRange("E1", "E5");

                // Union the three ranges into a single UnionRange
                UnionRange unionRange = range1.UnionRanges(new AsposeRange[] { range2, range3 });

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 10, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Use the union range as the data source for the chart
                chart.SetChartDataRange(unionRange.RefersTo, true);

                // Optional: set a title for clarity
                chart.Title.Text = "Non‑Sequential Data Chart";

                // Define output file path
                string outputPath = "NonSequentialChart.xlsx";

                // Ensure the directory exists (in case a relative path is used)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details for troubleshooting
                Console.Error.WriteLine($"Error: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}