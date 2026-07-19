// Title: C# – Set 3‑D Rotation and Elevation for a Column Chart using Aspose.Cells
// Description: This example builds a workbook, adds sample data, inserts a 3‑D column chart, and customizes its perspective by applying a 45° rotation and a 30° elevation before saving the file.
// Keywords: Aspose.Cells | C# | .NET | 3D column chart | RotationAngle | Elevation | chart perspective | Excel automation | chart styling | Aspose.Cells Chart API
// Common Searches: Aspose.Cells rotate 3D chart C# | set elevation angle Aspose.Cells .NET | Chart.RotationAngle property example | how to change 3D chart view Aspose.Cells | C# code for 3D column chart orientation | Aspose.Cells chart perspective settings
// Developer Intent: Programmatically adjust the viewing angle of a 3‑D column chart in an Excel workbook.
// Use Cases: Design a sales dashboard where the 3‑D column chart is tilted for visual emphasis. | Export financial data with a customized chart angle for presentation slides. | Automate report generation that requires consistent 3‑D chart orientation across workbooks.
// AI Prompts: Generate C# code that creates a 3‑D column chart with a 60° rotation and 20° elevation using Aspose.Cells. | Explain how RotationAngle and Elevation affect the appearance of a 3‑D chart in Aspose.Cells. | Show how to modify an existing chart's perspective properties after data binding in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example builds a workbook, adds sample data, inserts a 3‑D column chart, and customizes its perspective by applying a 45° rotation and a 30° elevation before saving the file.
    public class Enable3DRotationAndElevation
    {
        public static void Run()
        {
            try
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

                // Enable 3‑D rotation and elevation
                chart.RotationAngle = 45;
                chart.Elevation = 30;

                // Save the workbook
                string outputPath = "3DRotationElevationDemo.xlsx";
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
            Enable3DRotationAndElevation.Run();
        }
    }
}
