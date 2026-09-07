// Title: How to rotate primary Y‑axis tick labels 90° clockwise in an Aspose.Cells column chart using C#
// AI Prompts: Generate C# code that creates a column chart with Aspose.Cells and sets the primary Y‑axis tick labels to a 90‑degree clockwise rotation. | Show the steps to disable automatic rotation and apply a manual rotation angle to the value axis tick labels in an Aspose.Cells workbook. | Provide a complete example that saves a workbook after rotating the Y‑axis tick labels in a chart with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# rotate Y axis tick labels 90 degrees | set manual rotation angle for chart value axis labels Aspose.Cells .NET | disable automatic tick label rotation Aspose.Cells chart example | change orientation of primary Y‑axis labels in column chart using Aspose.Cells | Aspose.Cells column chart label rotation C# tutorial
// Tags: rotate Y‑axis tick labels Aspose.Cells C# | manual rotation angle chart value axis Aspose.Cells | disable automatic tick label rotation Aspose.Cells | column chart label orientation Aspose.Cells .NET | set tick label rotation angle Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new workbook, adds sample data, inserts a column chart, accesses the primary Y‑axis (value axis) tick labels, disables automatic rotation, sets the rotation angle to 90 degrees clockwise, and saves the workbook as RotatePrimaryYAxisTickLabels.xlsx.
    public class RotatePrimaryYAxisTickLabels
    {
        public static void Main(string[] args)
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Access the primary Y axis (value axis) tick labels
            TickLabels yAxisTickLabels = chart.ValueAxis.TickLabels;

            // Disable automatic rotation to apply manual angle
            yAxisTickLabels.IsAutomaticRotation = false;

            // Rotate the tick labels 90 degrees clockwise
            yAxisTickLabels.RotationAngle = 90;

            // Save the workbook to a file
            workbook.Save("RotatePrimaryYAxisTickLabels.xlsx");
        }
    }
}
