// Title: How to rotate X‑axis tick labels 45° in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and sets the X‑axis tick labels to a 45‑degree rotation. | Show how to disable automatic label rotation and apply a custom angle to the category axis in an Aspose.Cells workbook. | Provide a C# example that formats dense chart X‑axis labels by rotating them 45 degrees with Aspose.Cells.
// Common Searches: Aspose.Cells C# rotate category axis labels 45 degrees to avoid overlap | set custom tick label rotation for column chart in Aspose.Cells using C# | disable auto label rotation Aspose.Cells chart and specify angle | how to adjust X‑axis label angle in Aspose.Cells spreadsheet programmatically
// Tags: Aspose.Cells rotate X axis tick labels | category axis label rotation C# | disable automatic chart label rotation Aspose.Cells | column chart label formatting Aspose.Cells | Excel chart label angle Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, fills it with sample data, adds a column chart, assigns the data ranges, disables automatic label rotation, sets the X‑axis tick labels to a 45‑degree angle, and saves the file as RotateXAxisLabelsDemo.xlsx.
    public class RotateXAxisLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["A6"].PutValue("May");

                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(130);
                sheet.Cells["B6"].PutValue(170);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data source for the chart
                chart.NSeries.Add("B2:B6", true);          // Values
                chart.NSeries.CategoryData = "A2:A6";     // Categories (X‑axis)

                // Rotate X‑axis (category axis) tick labels by 45 degrees
                chart.CategoryAxis.TickLabels.IsAutomaticRotation = false; // disable auto‑rotation
                chart.CategoryAxis.TickLabels.RotationAngle = 45;

                // Save the workbook
                workbook.Save("RotateXAxisLabelsDemo.xlsx");
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
}
