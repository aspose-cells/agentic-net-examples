using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class HideHiddenDataPointsDemo
    {
        public static void Main()
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (Category in column A, Value in column B)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["B5"].PutValue(40);

            // Hide rows that contain data points we want to keep invisible in the chart
            // (e.g., hide rows 3 and 5)
            sheet.Cells.Rows[2].IsHidden = true; // hides row 3 (Category B)
            sheet.Cells.Rows[4].IsHidden = true; // hides row 5 (Category D)

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B2:B5", true);          // values
            chart.NSeries.CategoryData = "A2:A5";     // categories

            // Ensure that only visible cells are plotted.
            chart.PlotVisibleCellsOnly = true; // default is true

            // Define output file path
            string outputPath = "HideHiddenDataPointsDemo.xlsx";

            // Ensure output directory exists
            string directory = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
        }
    }
}