using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class ChartCloneAndCompare
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Q1");
            sheet.Cells["A3"].PutValue("Q2");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(200);

            // Add the original chart
            int originalChartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
            Chart originalChart = sheet.Charts[originalChartIndex];
            originalChart.NSeries.Add("B2:B3", true);
            originalChart.NSeries.CategoryData = "A2:A3";

            // Add the cloned chart and duplicate the series settings manually
            int clonedChartIndex = sheet.Charts.Add(ChartType.Column, 5, 6, 15, 11);
            Chart clonedChart = sheet.Charts[clonedChartIndex];
            clonedChart.NSeries.Add("B2:B3", true);
            clonedChart.NSeries.CategoryData = "A2:A3";

            // Change the legend entry text fill of the cloned chart to none
            LegendEntry clonedLegendEntry = clonedChart.NSeries[0].LegendEntry;
            clonedLegendEntry.IsTextNoFill = true;

            // Render both charts to images
            string originalImagePath = "OriginalChart.png";
            string clonedImagePath = "ClonedChart.png";
            originalChart.ToImage(originalImagePath);
            clonedChart.ToImage(clonedImagePath);

            // Verify that the image files were created before comparing
            if (!File.Exists(originalImagePath) || !File.Exists(clonedImagePath))
            {
                Console.WriteLine("One or both chart images were not generated.");
                return;
            }

            // Load the images as byte arrays for pixel‑wise (byte‑wise) comparison
            byte[] originalBytes = File.ReadAllBytes(originalImagePath);
            byte[] clonedBytes = File.ReadAllBytes(clonedImagePath);

            if (originalBytes.Length != clonedBytes.Length)
            {
                Console.WriteLine("Images have different sizes; cannot compare.");
            }
            else
            {
                int diffCount = 0;
                for (int i = 0; i < originalBytes.Length; i++)
                {
                    if (originalBytes[i] != clonedBytes[i])
                        diffCount++;
                }

                Console.WriteLine($"Number of differing bytes: {diffCount}");
                Console.WriteLine(diffCount == 0
                    ? "Charts are visually identical."
                    : "Charts differ visually.");
            }

            // Save the workbook with both charts
            string workbookPath = "ChartCloneComparison.xlsx";
            workbook.Save(workbookPath);
            Console.WriteLine($"Workbook saved to '{workbookPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}