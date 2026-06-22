using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace ProgressBarChartWithDataLabels
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // -------------------------------------------------
                // Sample data for a progress bar chart (stacked bar)
                // -------------------------------------------------
                // Category names
                sheet.Cells["A1"].PutValue("Task");
                sheet.Cells["A2"].PutValue("Task 1");
                sheet.Cells["A3"].PutValue("Task 2");
                sheet.Cells["A4"].PutValue("Task 3");

                // Completed portion (percentage)
                sheet.Cells["B1"].PutValue("Completed");
                sheet.Cells["B2"].PutValue(0.40); // 40%
                sheet.Cells["B3"].PutValue(0.70); // 70%
                sheet.Cells["B4"].PutValue(0.55); // 55%

                // Remaining portion (percentage)
                sheet.Cells["C1"].PutValue("Remaining");
                sheet.Cells["C2"].PutValue(0.60);
                sheet.Cells["C3"].PutValue(0.30);
                sheet.Cells["C4"].PutValue(0.45);

                // Add a stacked bar chart (used as a progress bar)
                int chartIndex = sheet.Charts.Add(ChartType.BarStacked, 6, 0, 20, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Add the two series (Completed and Remaining)
                chart.NSeries.Add("B2:B4", true); // Completed series
                chart.NSeries.Add("C2:C4", true); // Remaining series

                // Set category (X) axis data
                chart.NSeries.CategoryData = "A2:A4";

                // -------------------------------------------------
                // Add data labels to each series showing exact percentages
                // -------------------------------------------------
                foreach (Series series in chart.NSeries)
                {
                    // Enable data labels for the series
                    series.DataLabels.ShowPercentage = true;   // Show percentage value
                    series.DataLabels.ShowValue = false;      // Hide raw value
                    series.DataLabels.Position = LabelPositionType.InsideEnd; // Position inside the bar

                    // Format as percentage with two decimal places
                    series.DataLabels.NumberFormat = "0.00%";

                    // Make the label font bold for better readability
                    series.DataLabels.Font.IsBold = true;
                }

                // Save the workbook
                string outputPath = "ProgressBarChartWithDataLabels.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}