using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisGetAxisTextsErrorHandlingDemo
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

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Attempt to retrieve axis labels before calling Calculate()
            string[] axisLabels = null;
            try
            {
                // This call is expected to throw because Calculate() hasn't been executed yet
                axisLabels = chart.ValueAxis.GetAxisTexts();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error retrieving axis texts before calculation: " + ex.Message);
                // Perform chart calculation to generate axis labels
                chart.Calculate();
                // Retry after calculation
                axisLabels = chart.ValueAxis.GetAxisTexts();
            }

            // Output the retrieved axis labels
            Console.WriteLine("Value Axis Labels:");
            if (axisLabels != null)
            {
                foreach (string label in axisLabels)
                {
                    Console.WriteLine(label);
                }
            }

            // Save the workbook (lifecycle rule)
            workbook.Save("AxisGetAxisTextsErrorHandlingDemo.xlsx");
        }
    }
}