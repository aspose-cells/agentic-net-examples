using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data: numeric X values in column A and Y values in column B
            worksheet.Cells["A1"].PutValue("X");
            worksheet.Cells["B1"].PutValue("Y");
            for (int i = 2; i <= 6; i++)
            {
                worksheet.Cells[i, 0].PutValue(i - 1);          // X = 1,2,3,4,5
                worksheet.Cells[i, 1].PutValue((i - 1) * 10);   // Y = 10,20,30,40,50
            }

            // Add a column chart (category X axis by default)
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 10, 0, 25, 15);
            Chart chart = worksheet.Charts[chartIndex];
            chart.NSeries.Add("B2:B6", false);   // Y values
            chart.NSeries[0].Name = "Series1";

            // Determine if the chart already uses a value axis for X.
            // Only Scatter charts use a value axis for X in Aspose.Cells.
            bool isXValueAxis = chart.Type == ChartType.Scatter;

            if (!isXValueAxis)
            {
                // Convert to a Scatter chart to support numeric X values
                chart.Type = ChartType.Scatter;

                // Assign numeric X values to the series
                chart.NSeries[0].XValues = "A2:A6";

                // Ensure the X axis is treated as a numeric (value) axis
                chart.CategoryAxis.CategoryType = CategoryType.AutomaticScale;
                chart.CategoryAxis.IsLogarithmic = false;
            }

            // Save the workbook
            string outputPath = "ChartXAxisValueAxis.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}