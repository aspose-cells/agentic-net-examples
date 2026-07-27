using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

public class RotateXAxisLabelsDemo
{
    public static void Run()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["A2"].PutValue("January");
            sheet.Cells["A3"].PutValue("February");
            sheet.Cells["A4"].PutValue("March");
            sheet.Cells["A5"].PutValue("April");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(180);
            sheet.Cells["B5"].PutValue(200);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 12);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart data source
            chart.NSeries.Add("B2:B5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Disable automatic rotation and rotate X‑axis (category axis) labels by 45 degrees
            chart.CategoryAxis.TickLabels.IsAutomaticRotation = false;
            chart.CategoryAxis.TickLabels.RotationAngle = 45;

            // Ensure output directory exists
            string outputPath = "RotatedXAxisLabels.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}

public class Program
{
    public static void Main(string[] args)
    {
        RotateXAxisLabelsDemo.Run();
    }
}