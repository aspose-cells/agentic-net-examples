// Title: Aspose.Cells for .NET – Set Every Chart Background to Light Gray
// Description: Loads a workbook, iterates through each worksheet and its charts, applies a light‑gray fill to the chart area using ChartArea.Area.BackgroundColor, and saves the updated file.
// Keywords: Aspose.Cells | C# chart background | set chart area color | loop through charts | light gray Excel chart | Excel workbook formatting | ChartArea.BackgroundColor
// Common Searches: change background color of all charts Aspose.Cells .NET | loop through worksheets and charts to set color | Aspose.Cells set chart area fill | apply uniform chart background in Excel with C# | Aspose.Cells chart formatting example
// Developer Intent: Apply a light‑gray background to every chart in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Standardize chart appearance in monthly financial reports. | Enforce corporate branding colors across template workbooks. | Ensure visual consistency when generating charts from automated data pipelines.
// AI Prompts: Generate C# code with Aspose.Cells that sets the background color of all charts in a workbook to a specified shade and saves the file. | Show how to iterate over worksheets and charts in Aspose.Cells and modify ChartArea.Area.BackgroundColor. | Explain the purpose of ChartArea.Area.BackgroundColor and how to use it for uniform chart styling.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads a workbook, iterates through each worksheet and its charts, applies a light‑gray fill to the chart area using ChartArea.Area.BackgroundColor, and saves the updated file.
class SetChartBackgroundColor
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
                // Set the background color of the chart area to light gray
                chart.ChartArea.Area.BackgroundColor = Color.LightGray;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
