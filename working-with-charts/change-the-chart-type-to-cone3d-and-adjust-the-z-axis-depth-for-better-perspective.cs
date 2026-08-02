// Title: Create a 3‑D cone‑style chart with Aspose.Cells for .NET – adjust depth and perspective
// Description: This example builds a new Workbook, fills cells A1:B4 with category and sales data, adds a 3‑D Column chart (used to emulate a Cone3D chart), assigns the series, sets DepthPercent to 250, applies a 30‑degree perspective, disables RightAngleAxes, and saves the file as Cone3DChart.xlsx.
// Keywords: Aspose.Cells | .NET | 3D chart | Cone3D | Column3D | DepthPercent | Perspective | RightAngleAxes | Excel chart customization | chart type conversion
// Common Searches: Aspose.Cells create Cone3D chart .NET | how to set chart depth percent Aspose.Cells | adjust perspective of 3D chart Aspose.Cells | simulate cone chart with Column3D Aspose.Cells | change RightAngleAxes property Aspose.Cells
// Developer Intent: Generate a 3‑D cone‑shaped chart and enhance its visual depth using Aspose.Cells for .NET.
// Use Cases: Produce a sales dashboard where figures are displayed in a cone‑style 3‑D chart for greater visual impact. | Retrofit existing 3‑D charts with deeper perspective to improve readability in financial reports. | Create presentation‑ready Excel files that require a cone‑like appearance without native Cone3D support.
// AI Prompts: Write C# code with Aspose.Cells to emulate a Cone3D chart by using Column3D, set DepthPercent to 250, and apply a 30‑degree perspective. | Explain how to increase the Z‑axis depth of a 3‑D chart in Aspose.Cells and why disabling RightAngleAxes is necessary. | Show how to convert any 3‑D column chart to a cone‑style visual using Aspose.Cells properties.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// This example builds a new Workbook, fills cells A1:B4 with category and sales data, adds a 3‑D Column chart (used to emulate a Cone3D chart), assigns the series, sets DepthPercent to 250, applies a 30‑degree perspective, disables RightAngleAxes, and saves the file as Cone3DChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            worksheet.Cells["A1"].PutValue("Category");
            worksheet.Cells["A2"].PutValue("Q1");
            worksheet.Cells["A3"].PutValue("Q2");
            worksheet.Cells["A4"].PutValue("Q3");
            worksheet.Cells["B1"].PutValue("Sales");
            worksheet.Cells["B2"].PutValue(1000);
            worksheet.Cells["B3"].PutValue(2000);
            worksheet.Cells["B4"].PutValue(3000);

            // Add a 3‑D column chart (Cone3D is not available in this Aspose.Cells version)
            int chartIndex = worksheet.Charts.Add(ChartType.Column3D, 5, 0, 20, 8);
            Chart chart = worksheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Adjust the Z‑axis depth (percentage of chart width) for a stronger 3‑D effect
            chart.DepthPercent = 250; // 250 % depth

            // Optional: improve perspective visibility
            chart.Perspective = 30;          // Perspective angle (0‑100)
            chart.RightAngleAxes = false;    // Ensure perspective is applied

            // Save the workbook with the chart
            string outputPath = "Cone3DChart.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
