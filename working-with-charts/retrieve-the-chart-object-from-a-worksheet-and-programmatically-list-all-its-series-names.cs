using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartSeriesNames
{
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");

            worksheet.Cells["B1"].PutValue("Series 1");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            worksheet.Cells["C1"].PutValue("Series 2");
            worksheet.Cells["C2"].PutValue(15);
            worksheet.Cells["C3"].PutValue(25);
            worksheet.Cells["C4"].PutValue(35);

            // Add a chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data source for the chart series
            chart.NSeries.Add("B2:B4", true); // Series 1 values
            chart.NSeries.Add("C2:C4", true); // Series 2 values
            chart.NSeries.CategoryData = "A2:A4";

            // Optionally assign names to the series (referencing cells)
            chart.NSeries[0].Name = "=A1"; // "Series 1"
            chart.NSeries[1].Name = "=C1"; // "Series 2"

            // Retrieve the chart object (already have it) and list all series names
            Console.WriteLine("Series names in the chart:");
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];

                // Use the Name property if set; otherwise fall back to DisplayName
                string seriesName = !string.IsNullOrEmpty(series.Name) ? series.Name : series.DisplayName;
                Console.WriteLine($"Series {i + 1}: {seriesName}");
            }

            // Save the workbook (output file name can be changed as needed)
            workbook.Save("ChartSeriesNamesOutput.xlsx");
        }
    }
}