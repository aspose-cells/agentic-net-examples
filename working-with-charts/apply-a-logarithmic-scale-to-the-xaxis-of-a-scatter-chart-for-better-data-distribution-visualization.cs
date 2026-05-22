using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsLogScaleDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and obtain the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data: exponential X values and linear Y values
            sheet.Cells["A1"].PutValue("X Values");
            sheet.Cells["B1"].PutValue("Y Values");
            for (int i = 2; i <= 11; i++)
            {
                // X = 2^(i-2) => 1, 2, 4, 8, ...
                sheet.Cells[$"A{i}"].PutValue(Math.Pow(2, i - 2));
                sheet.Cells[$"B{i}"].PutValue(i * 5);
            }

            // Insert a scatter chart
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Bind Y values and X values to the first series
            chart.NSeries.Add("B2:B11", true);          // Y values range
            chart.NSeries[0].XValues = "A2:A11";        // X values range
            chart.NSeries[0].Name = "Exponential Data";

            // Apply logarithmic scaling to the X (category) axis
            chart.CategoryAxis.IsLogarithmic = true;    // Axis.IsLogarithmic property
            chart.CategoryAxis.LogBase = 2;             // Axis.LogBase property (base 2)

            // Optional: set a title for the X axis
            chart.CategoryAxis.Title.Text = "Logarithmic X Axis (Base 2)";

            // Save the workbook with the chart
            workbook.Save("ScatterLogScale.xlsx", SaveFormat.Xlsx);
        }
    }
}