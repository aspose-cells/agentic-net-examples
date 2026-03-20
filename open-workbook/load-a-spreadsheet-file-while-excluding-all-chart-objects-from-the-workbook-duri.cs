using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartExclusionDemo
{
    class Program
    {
        static void Main()
        {
            // Prepare a sample workbook with a chart
            string sourceFile = "InputWithCharts.xlsx";
            CreateSampleWorkbook(sourceFile);

            // LoadOptions with a LoadFilter that excludes charts
            LoadOptions loadOptions = new LoadOptions
            {
                LoadFilter = new LoadFilter(LoadDataFilterOptions.All & ~LoadDataFilterOptions.Chart)
            };

            // Load the workbook using the specified options
            Workbook workbook = new Workbook(sourceFile, loadOptions);

            // Verify that charts are not loaded
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet '{sheet.Name}' contains {sheet.Charts.Count} chart(s).");
            }

            // Save the workbook without charts
            string outputFile = "OutputWithoutCharts.xlsx";
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved without charts to '{outputFile}'.");
        }

        private static void CreateSampleWorkbook(string filePath)
        {
            // Create a new workbook and add some data
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            ws.Name = "DataSheet";

            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(30);

            // Add a column chart
            int chartIndex = ws.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = ws.Charts[chartIndex];
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";
            chart.Title.Text = "Sample Chart";

            // Save the workbook
            wb.Save(filePath, SaveFormat.Xlsx);
        }
    }
}