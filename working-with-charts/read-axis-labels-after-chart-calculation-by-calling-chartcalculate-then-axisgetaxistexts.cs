using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisLabelsDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(8000);
            worksheet.Cells["B3"].PutValue(4000);
            worksheet.Cells["B4"].PutValue(-8000);

            // Add a column chart to the worksheet
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the series and the category axis
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Calculate the chart to generate axis labels
            chart.Calculate();

            // Retrieve the axis labels using the GetAxisTexts method
            string[] valueAxisLabels = chart.ValueAxis.GetAxisTexts();
            string[] categoryAxisLabels = chart.CategoryAxis.GetAxisTexts();

            // Output the retrieved labels
            Console.WriteLine("Value Axis Labels:");
            foreach (string label in valueAxisLabels)
            {
                Console.WriteLine(label);
            }

            Console.WriteLine("\nCategory Axis Labels:");
            foreach (string label in categoryAxisLabels)
            {
                Console.WriteLine(label);
            }

            // Save the workbook (optional, demonstrates lifecycle usage)
            workbook.Save("AxisLabelsDemo.xlsx");
        }
    }
}