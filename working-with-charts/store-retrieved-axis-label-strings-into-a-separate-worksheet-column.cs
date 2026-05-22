using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class StoreAxisLabelsInWorksheet
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Populate sample data for the chart (categories and values)
                worksheet.Cells["A1"].PutValue("Category");
                worksheet.Cells["B1"].PutValue("Value");
                worksheet.Cells["A2"].PutValue("Alpha");
                worksheet.Cells["B2"].PutValue(15);
                worksheet.Cells["A3"].PutValue("Beta");
                worksheet.Cells["B3"].PutValue(30);
                worksheet.Cells["A4"].PutValue("Gamma");
                worksheet.Cells["B4"].PutValue(45);

                // Add a column chart to the worksheet
                int chartIndex = worksheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
                Chart chart = worksheet.Charts[chartIndex];

                // Set the data range for the series and the category axis
                chart.NSeries.Add("B2:B4", true);          // Values
                chart.NSeries.CategoryData = "A2:A4";      // Categories

                // Calculate the chart so that axis labels are generated
                chart.Calculate();

                // Retrieve the category axis labels
                string[] axisLabels = chart.CategoryAxis.GetAxisTexts();

                // Store the retrieved labels into column C, starting from row 2
                int startRow = 1; // zero‑based index (row 2 in Excel)
                for (int i = 0; i < axisLabels.Length; i++)
                {
                    // Cells[row, column] uses zero‑based indices; column 2 corresponds to "C"
                    worksheet.Cells[startRow + i, 2].PutValue(axisLabels[i]);
                }

                // Add a header for the stored labels
                worksheet.Cells["C1"].PutValue("Axis Labels");

                // Save the workbook (ensure the directory exists)
                string outputPath = "AxisLabelsStored.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Runtime error: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            try
            {
                StoreAxisLabelsInWorksheet.Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unhandled exception: {ex.Message}");
            }
        }
    }
}