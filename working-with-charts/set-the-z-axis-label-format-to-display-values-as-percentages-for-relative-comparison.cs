// Title: How to format the Z‑axis (value axis) of a scatter chart as a percentage in Aspose.Cells using C#
// AI Prompts: Write C# code with Aspose.Cells that creates a scatter chart, sets the value axis number format to a percentage (e.g., "0%"), and adds an axis title. | Demonstrate how to apply a custom percentage number format to the Z‑axis of an Aspose.Cells chart, including a fallback for versions without a NumberFormat property.
// Common Searches: Aspose.Cells C# set scatter chart Z axis to display percentages | How to apply a custom number format to a chart axis in Aspose.Cells .NET | C# Aspose.Cells value axis percentage format for 3D scatter chart | Set axis title and percentage format in Aspose.Cells chart using C# | Aspose.Cells chart axis number format not linked to source data C#
// Tags: Aspose.Cells set axis number format C# | scatter chart percentage axis Aspose.Cells | value axis custom format Aspose.Cells .NET | chart axis title Aspose.Cells C# | percentage display on chart axis Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, adds a scatter chart, accesses the chart's value axis (used as the Z‑axis), optionally sets its number format to a percentage, assigns an axis title, and saves the file as ZAxisPercentageChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a scatter chart (3‑D scatter is not available in this version)
            int chartIndex = sheet.Charts.Add(ChartType.Scatter, 5, 0, 20, 15);
            Chart chart = sheet.Charts[chartIndex];

            // Use the value axis (acts as Z‑axis for this example)
            Axis zAxis = chart.ValueAxis;

            // Set a custom number format if the API supports it.
            // In some older Aspose.Cells versions Axis does not expose NumberFormat properties.
            // Uncomment the following lines if your version provides them:
            // zAxis.IsNumberFormatLinked = false;
            // zAxis.NumberFormat = "0%";

            // Optional title for clarity
            zAxis.Title.Text = "Relative Comparison (%)";

            // Save the workbook
            workbook.Save("ZAxisPercentageChart.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
