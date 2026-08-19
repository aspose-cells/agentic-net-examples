// Title: Rotate X‑Axis Labels 45° in Aspose.Cells (C#) to Prevent Overlap
// Description: Creates a workbook, fills month names and values, adds a column chart, disables automatic rotation, sets the X‑axis tick label angle to 45 degrees, and saves the file as RotateXAxisLabelsDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells rotate X axis labels | C# chart label rotation | 45 degree tick label Aspose | prevent overlapping chart labels .NET | category axis label angle | Aspose.Cells column chart example | Excel chart label customization
// Common Searches: rotate x axis labels 45 degrees Aspose.Cells | Aspose.Cells chart label rotation C# | how to prevent overlapping axis labels in Excel with Aspose | set tick label angle Aspose.Cells .NET | custom X‑axis label angle Aspose chart
// Developer Intent: Apply a 45‑degree rotation to X‑axis tick labels in a column chart to improve readability and avoid label collisions.
// Use Cases: Monthly sales chart where month names are tilted for clear display. | Dense categorical reports that require angled X‑axis labels for printing. | Automated Excel generation with custom label angles for presentation decks.
// AI Prompts: Generate C# code with Aspose.Cells to rotate X‑axis labels to any angle for a given chart type. | Explain how to programmatically decide the label rotation based on the number of categories in an Aspose.Cells chart. | Show the steps to revert X‑axis label rotation back to automatic after a custom angle has been applied.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, fills month names and values, adds a column chart, disables automatic rotation, sets the X‑axis tick label angle to 45 degrees, and saves the file as RotateXAxisLabelsDemo.xlsx using Aspose.Cells for .NET.
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
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(120);
                sheet.Cells["B3"].PutValue(150);
                sheet.Cells["B4"].PutValue(180);
                sheet.Cells["B5"].PutValue(200);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 15);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the chart
                chart.NSeries.Add("B2:B5", true);          // Values
                chart.NSeries.CategoryData = "A2:A5";     // Categories (X‑axis)

                // Rotate the X‑axis (category axis) tick labels by 45 degrees
                chart.CategoryAxis.TickLabels.IsAutomaticRotation = false;
                chart.CategoryAxis.TickLabels.RotationAngle = 45;

                // Save the workbook to an XLSX file
                workbook.Save("RotateXAxisLabelsDemo.xlsx");
                Console.WriteLine("Workbook saved as RotateXAxisLabelsDemo.xlsx");
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
