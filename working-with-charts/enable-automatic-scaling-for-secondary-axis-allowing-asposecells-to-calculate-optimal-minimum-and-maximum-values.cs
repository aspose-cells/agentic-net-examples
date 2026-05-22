using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsDemo
{
    class SecondaryAxisAutoScalingDemo
    {
        static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook created successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["B3"].PutValue(200);
            sheet.Cells["B4"].PutValue(300);

            sheet.Cells["C1"].PutValue("Series 2");
            sheet.Cells["C2"].PutValue(5000);
            sheet.Cells["C3"].PutValue(3000);
            sheet.Cells["C4"].PutValue(1000);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Add two series and set category data
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.Add("C2:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Plot the second series on the secondary value axis
            chart.NSeries[1].PlotOnSecondAxis = true;

            // Access the secondary value axis and enable automatic scaling
            Axis secondaryAxis = chart.SecondValueAxis;
            secondaryAxis.IsAutomaticMinValue = true; // Let Aspose.Cells calculate optimal minimum
            secondaryAxis.IsAutomaticMaxValue = true; // Let Aspose.Cells calculate optimal maximum
            secondaryAxis.Title.Text = "Secondary Axis";

            // Define output file path
            string outputPath = "SecondaryAxisAutoScalingDemo.xlsx";

            // Ensure the directory exists (optional safety)
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath);
        }
    }
}