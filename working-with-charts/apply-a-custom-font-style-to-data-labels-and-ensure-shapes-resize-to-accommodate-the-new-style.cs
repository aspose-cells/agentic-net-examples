using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    public class CustomDataLabelFontAndResizeDemo
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
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("A");
            worksheet.Cells["A3"].PutValue("B");
            worksheet.Cells["A4"].PutValue("C");
            worksheet.Cells["A5"].PutValue("D");

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);
            worksheet.Cells["B5"].PutValue(40);

            // Add a column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Enable data labels for the first series
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;

            // Apply custom font style to the data labels
            series.DataLabels.Font.Name = "Calibri";
            series.DataLabels.Font.Size = 16;
            series.DataLabels.Font.Color = Color.Blue;
            series.DataLabels.Font.IsBold = true;

            // Ensure the shape of each data label resizes to fit the new font
            series.DataLabels.IsResizeShapeToFitText = true;
            series.DataLabels.AutoScaleFont = true;

            // Apply the font settings to all child nodes of the data labels
            series.DataLabels.ApplyFont();

            // Enforce resizing for individual point labels as well
            foreach (ChartPoint point in series.Points)
            {
                point.DataLabels.IsResizeShapeToFitText = true;
                point.DataLabels.AutoScaleFont = true;
                point.DataLabels.Font.Name = "Calibri";
                point.DataLabels.Font.Size = 16;
                point.DataLabels.Font.Color = Color.Blue;
                point.DataLabels.Font.IsBold = true;
                point.DataLabels.ApplyFont();
            }

            // Define output file path
            string outputPath = "CustomDataLabelFontAndResizeDemo.xlsx";

            // Save the workbook (overwrite if exists)
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
    }
}