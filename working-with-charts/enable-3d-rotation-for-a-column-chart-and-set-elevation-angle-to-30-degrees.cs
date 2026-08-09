// Title: Aspose.Cells for .NET – Enable 3‑D Rotation & Set Elevation (30°) on a Column Chart (C#)
// Description: Creates a new workbook, adds sample data, inserts a 3‑D column chart, and configures its perspective by applying a 45° rotation and a 30° elevation before saving the file.
// Keywords: Aspose.Cells 3D chart rotation C# | Chart.Elevation property Aspose.Cells | Column3D chart Aspose.Cells example | set RotationAngle Aspose.Cells .NET | 3‑D column chart visualization | Aspose.Cells chart view angles
// Common Searches: how to rotate a 3‑D column chart using Aspose.Cells | set elevation angle for 3‑D chart Aspose.Cells .NET | Aspose.Cells Chart.RotationAngle C# example | adjust 3‑D view angles in Excel with Aspose.Cells | C# code for 3‑D column chart rotation and elevation
// Developer Intent: Add a 3‑D column chart to a workbook and programmatically set its rotation and elevation angles.
// Use Cases: Produce a sales dashboard where the 3‑D column chart is tilted for clearer data comparison. | Generate presentation‑ready Excel reports with a consistent 45° rotation and 30° elevation across all charts. | Automate bulk styling of existing 3‑D charts in multiple workbooks by updating their view angles via code.
// AI Prompts: Write C# code with Aspose.Cells that creates a 3‑D column chart, sets RotationAngle to 45 degrees, Elevation to 30 degrees, and saves the workbook. | Explain the impact of RotationAngle and Elevation on the visual rendering of a 3‑D chart in Aspose.Cells. | Provide step‑by‑step instructions to modify an existing column chart's 3‑D rotation and elevation using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a new workbook, adds sample data, inserts a 3‑D column chart, and configures its perspective by applying a 45° rotation and a 30° elevation before saving the file.
    public class Enable3DRotationAndSetElevation
    {
        public static void Main(string[] args)
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

            worksheet.Cells["B1"].PutValue("Value");
            worksheet.Cells["B2"].PutValue(10);
            worksheet.Cells["B3"].PutValue(20);
            worksheet.Cells["B4"].PutValue(30);

            // Add a 3‑D column chart
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable 3‑D rotation by setting the RotationAngle property
            chart.RotationAngle = 45; // rotation around the Z‑axis in degrees

            // Set the elevation angle to 30 degrees
            chart.Elevation = 30; // view height angle in degrees

            // Save the workbook to a file
            string outputPath = "3DRotation_Elevation.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine($"3‑D column chart created with rotation and elevation set. Saved to '{outputPath}'.");
        }
    }
}
