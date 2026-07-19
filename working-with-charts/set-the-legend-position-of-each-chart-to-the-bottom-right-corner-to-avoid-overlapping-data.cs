// Title: Aspose.Cells C# – Set All Chart Legends to Bottom‑Right Corner and Disable Overlay
// Description: Loads a workbook, iterates through every worksheet and chart, sets each legend to the bottom‑right corner (LegendPositionType.Corner), turns off overlay (IsOverLay = false), and saves the updated file.
// Keywords: Aspose.Cells | C# | chart legend position | LegendPositionType.Corner | disable legend overlay | iterate worksheets | modify charts | Excel automation | .NET chart formatting
// Common Searches: Aspose.Cells set chart legend bottom right | C# move chart legend to corner Aspose | disable chart legend overlay Aspose.Cells | loop through charts in workbook C# | change legend position for all charts Aspose
// Developer Intent: Place each chart’s legend in the bottom‑right corner and prevent it from covering the chart area.
// Use Cases: Standardize legend placement across multiple charts before publishing a report. | Prepare an Excel file for printing where legends must not obscure data series. | Automate chart formatting in a template workbook to meet corporate style guidelines.
// AI Prompts: Generate C# code using Aspose.Cells to set all chart legends to the top‑left corner and enable overlay. | Provide an example that saves the workbook to a memory stream after adjusting legend positions. | Explain how to set legend position conditionally based on chart type in Aspose.Cells.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates through every worksheet and chart, sets each legend to the bottom‑right corner (LegendPositionType.Corner), turns off overlay (IsOverLay = false), and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all charts in the current worksheet
            foreach (Chart chart in sheet.Charts)
            {
                // Set the legend position to the bottom‑right corner
                chart.Legend.Position = LegendPositionType.Corner;

                // Ensure the legend does not overlay the chart area
                chart.Legend.IsOverLay = false;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
