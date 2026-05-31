using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class ConfigureYAxisAsValue
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data (numeric values) for the chart
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["A2"].PutValue("A");
                worksheet.Cells["A3"].PutValue("B");
                worksheet.Cells["A4"].PutValue("C");

                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["B2"].PutValue(120);
                worksheet.Cells["B3"].PutValue(250);
                worksheet.Cells["B4"].PutValue(370);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Configure the Y‑axis (ValueAxis) to represent numeric measurements accurately
                chart.ValueAxis.IsLogarithmic = false;               // Linear scale
                chart.ValueAxis.DisplayUnit = DisplayUnitType.None; // No scaling/division
                chart.ValueAxis.Title.Text = "Numeric Value";        // Optional descriptive title

                // Save the workbook to a file
                string outputPath = "ConfigureYAxisAsValue.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Application entry point
    public class Program
    {
        public static void Main(string[] args)
        {
            ConfigureYAxisAsValue.Run();
        }
    }
}