using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class AutomaticDisplayUnitDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (large values to trigger automatic display units)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(1500000);   // 1.5 million
                worksheet.Cells["B3"].PutValue(3000000);   // 3 million
                worksheet.Cells["B4"].PutValue(4500000);   // 4.5 million

                // Add a column chart
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable automatic scaling for the value (Y) axis
                Axis valueAxis = chart.ValueAxis;

                // Let Excel decide the best display unit (default is None)
                valueAxis.DisplayUnit = DisplayUnitType.None;

                // Show the display unit label (e.g., "Millions") if Excel selects one
                valueAxis.IsDisplayUnitLabelShown = true;

                // Ensure axis limits and major/minor units are automatic
                valueAxis.IsAutomaticMinValue = true;
                valueAxis.IsAutomaticMaxValue = true;
                valueAxis.IsAutomaticMajorUnit = true;
                valueAxis.IsAutomaticMinorUnit = true;

                // Determine output file path
                string outputPath = "AutomaticDisplayUnitDemo.xlsx";

                // Save the workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutomaticDisplayUnitDemo.Run();
        }
    }
}