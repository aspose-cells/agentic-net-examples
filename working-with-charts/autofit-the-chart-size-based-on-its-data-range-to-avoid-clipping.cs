using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartAutoFitDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(200);
            worksheet.Cells["B4"].PutValue(1500); // Larger value to test auto‑fit

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Ensure the chart layout is calculated (required before GetActualSize)
            chart.Calculate();

            // Retrieve the actual size needed to display the chart without clipping
            int[] actualSize = chart.GetActualSize(); // [0] = width, [1] = height

            // Apply the calculated size to the chart object
            chart.ChartObject.Width = actualSize[0];
            chart.ChartObject.Height = actualSize[1];

            // Save the workbook (save rule)
            workbook.Save("ChartAutoFitDemo.xlsx", SaveFormat.Xlsx);
        }
    }
}