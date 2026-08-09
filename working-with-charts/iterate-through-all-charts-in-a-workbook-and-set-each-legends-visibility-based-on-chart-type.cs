// Title: C# – Hide or Show Chart Legends by Type Across All Worksheets with Aspose.Cells
// Description: Loads an Excel workbook, iterates through every worksheet and each chart, detects pie‑related chart types (Pie, Pie3D, Doughnut, PieExploded, Pie3DExploded) and sets the ShowLegend property accordingly, then saves the file.
// Keywords: Aspose.Cells chart legend visibility | C# hide chart legend | iterate charts workbook Aspose | Chart.ShowLegend property | pie chart legend Aspose.Cells | .NET Excel chart automation
// Common Searches: how to hide legends for pie charts using Aspose.Cells C# | loop through all charts in an Excel workbook and set legend visibility | Aspose.Cells ShowLegend based on chart type | C# programmatically hide chart legends in Excel | Aspose.Cells iterate worksheets charts
// Developer Intent: Loop through every chart in an Excel workbook and toggle the legend visibility depending on whether the chart is a pie‑type.
// Use Cases: Remove legends from pie, doughnut, and exploded pie charts while keeping them for column or line charts in a financial dashboard. | Standardize legend settings across multiple sheets before exporting the workbook to PDF or image formats. | Prepare a printable report by suppressing unnecessary legends on pie charts to reduce visual clutter.
// AI Prompts: Write C# code with Aspose.Cells that iterates all charts in a workbook and hides legends for pie‑related types. | Show how to check ChartType and set the ShowLegend property for each chart in a .NET Excel file. | Explain how to extend the legend‑visibility logic to include additional chart types such as radar, bubble, or scatter.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLegendVisibility
{
    // Loads an Excel workbook, iterates through every worksheet and each chart, detects pie‑related chart types (Pie, Pie3D, Doughnut, PieExploded, Pie3DExploded) and sets the ShowLegend property accordingly, then saves the file.
    class Program
    {
        static void Main()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Iterate through each chart on the worksheet
                    foreach (Chart chart in sheet.Charts)
                    {
                        // Hide legends for pie‑related charts; show for all other types
                        bool hideLegend = chart.Type == ChartType.Pie ||
                                          chart.Type == ChartType.Pie3D ||
                                          chart.Type == ChartType.Doughnut ||
                                          chart.Type == ChartType.PieExploded ||
                                          chart.Type == ChartType.Pie3DExploded;

                        // Apply legend visibility
                        chart.ShowLegend = !hideLegend;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
