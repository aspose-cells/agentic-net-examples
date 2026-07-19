// Title: Rotate X‑Axis Labels 45° in Aspose.Cells .NET Column Chart to Avoid Overlap
// Description: Creates a workbook, fills cells A1:B5 with month and sales data, adds a column chart, links the data range, disables automatic label rotation, sets a 45‑degree manual rotation for the category axis tick labels, and saves the file as RotateXAxisLabelsDemo.xlsx.
// Keywords: Aspose.Cells rotate X axis labels | C# chart label rotation | CategoryAxis.TickLabels RotationAngle | prevent overlapping chart labels | manual tick label angle Aspose.Cells | .NET Excel chart formatting
// Common Searches: rotate X axis labels 45 degrees Aspose.Cells .NET | Aspose.Cells chart tick label rotation example | how to prevent overlapping X axis labels in Excel chart using C# | set manual rotation for chart category axis labels Aspose | Aspose.Cells column chart label angle
// Developer Intent: Apply a 45° manual rotation to X‑axis (category) tick labels in a column chart to keep labels readable.
// Use Cases: Display monthly sales in a column chart where month names would otherwise clash. | Generate Excel reports with multiple charts that require consistent label angles for printing. | Create dashboards where narrow category columns need angled labels for better visual clarity.
// AI Prompts: Show C# code to rotate X‑axis tick labels 45 degrees in an Aspose.Cells column chart and disable auto‑rotation. | Give an example of setting a custom RotationAngle for CategoryAxis.TickLabels in Aspose.Cells and saving the workbook. | Explain how to adjust label rotation for different chart types (column, line, bar) using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills cells A1:B5 with month and sales data, adds a column chart, links the data range, disables automatic label rotation, sets a 45‑degree manual rotation for the category axis tick labels, and saves the file as RotateXAxisLabelsDemo.xlsx.
    public class RotateXAxisLabelsDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B1"].PutValue("Sales");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories (X‑axis)

                // Rotate the X‑axis (category axis) tick labels by 45 degrees
                chart.CategoryAxis.TickLabels.IsAutomaticRotation = false; // Disable auto‑rotation
                chart.CategoryAxis.TickLabels.RotationAngle = 45;          // Set manual rotation

                // Save the workbook to a file
                string outputPath = "RotateXAxisLabelsDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            RotateXAxisLabelsDemo.Run();
        }
    }
}
