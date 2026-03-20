using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSkipDemo
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input_with_charts.xlsx";

            if (!File.Exists(sourcePath))
            {
                var tempWb = new Workbook();
                var ws = tempWb.Worksheets[0];
                ws.Cells["A1"].PutValue(10);
                ws.Cells["A2"].PutValue(20);
                ws.Cells["A3"].PutValue(30);
                int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = ws.Charts[chartIdx];
                chart.NSeries.Add("A1:A3", true);
                tempWb.Save(sourcePath, SaveFormat.Xlsx);
            }

            LoadOptions loadOptions = new LoadOptions();
            LoadFilter filter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart);
            loadOptions.LoadFilter = filter;
            loadOptions.KeepUnparsedData = false;

            Workbook workbook = new Workbook(sourcePath, loadOptions);

            Console.WriteLine("Worksheets loaded: " + workbook.Worksheets.Count);

            string outputPath = "output_without_charts.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook saved without chart data to: " + outputPath);
        }
    }
}