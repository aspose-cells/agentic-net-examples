using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartVerification
{
    class Program
    {
        static void Main()
        {
            // Paths to source workbooks
            string[] sourcePaths = { "Source1.xlsx", "Source2.xlsx" };

            // Ensure source workbooks exist; create simple ones with a chart if missing
            foreach (string path in sourcePaths)
            {
                if (!File.Exists(path))
                {
                    CreateSampleWorkbook(path);
                }
            }

            // Create the destination workbook (empty)
            Workbook destWorkbook = new Workbook();

            // Store chart information from each source workbook before combining
            var sourceChartInfo = new List<(string SheetName, int ChartIndex, ChartType Type)>();

            // Load each source workbook, record its charts, then combine into destination
            foreach (string srcPath in sourcePaths)
            {
                // Load source workbook
                Workbook srcWorkbook = new Workbook(srcPath);

                // Record chart details from all worksheets
                foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                {
                    for (int i = 0; i < srcSheet.Charts.Count; i++)
                    {
                        Chart srcChart = srcSheet.Charts[i];
                        sourceChartInfo.Add((srcSheet.Name, i, srcChart.Type));
                    }
                }

                // Combine the source workbook into the destination workbook
                destWorkbook.Combine(srcWorkbook);
            }

            // Refresh all charts/pivots after combination
            destWorkbook.Worksheets.RefreshAll();

            // Verify that each recorded chart exists in the combined workbook
            bool allChartsPresent = true;
            foreach (var info in sourceChartInfo)
            {
                // Find the worksheet by name (combined workbook keeps original sheet names)
                Worksheet combinedSheet = destWorkbook.Worksheets[info.SheetName];
                if (combinedSheet == null)
                {
                    Console.WriteLine($"Missing worksheet: {info.SheetName}");
                    allChartsPresent = false;
                    continue;
                }

                // Check chart count
                if (info.ChartIndex >= combinedSheet.Charts.Count)
                {
                    Console.WriteLine($"Worksheet '{info.SheetName}' is missing chart at index {info.ChartIndex}");
                    allChartsPresent = false;
                    continue;
                }

                // Verify chart type matches
                Chart combinedChart = combinedSheet.Charts[info.ChartIndex];
                if (combinedChart.Type != info.Type)
                {
                    Console.WriteLine($"Chart type mismatch in worksheet '{info.SheetName}' at index {info.ChartIndex}. Expected: {info.Type}, Found: {combinedChart.Type}");
                    allChartsPresent = false;
                }
            }

            Console.WriteLine(allChartsPresent
                ? "All charts from source workbooks are present in the combined workbook."
                : "Some charts are missing or mismatched in the combined workbook.");

            // Save the combined workbook
            destWorkbook.Save("Combined.xlsx", SaveFormat.Xlsx);
        }

        private static void CreateSampleWorkbook(string filePath)
        {
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = Path.GetFileNameWithoutExtension(filePath);

            // Add sample data
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 7);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            wb.Save(filePath, SaveFormat.Xlsx);
        }
    }
}