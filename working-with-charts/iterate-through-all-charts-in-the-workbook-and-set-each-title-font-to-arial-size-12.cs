// Title: Set all chart titles to Arial 12‑point font in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file, iterates through every worksheet and chart, makes each chart title visible, and sets the title font to Arial 12 pt using Aspose.Cells. | Write a .NET script to bulk‑apply a specific font name and size to chart titles across all sheets in a workbook with Aspose.Cells. | Create a function that updates the Font.Name and Font.Size properties of Chart.Title for every chart in a workbook and saves the result.
// Common Searches: Aspose.Cells C# change font of all chart titles in workbook | How to set chart title font to Arial 12pt programmatically with Aspose.Cells | Iterate over worksheets and charts to modify title font size in .NET Excel file | Bulk update chart title styling using Aspose.Cells API
// Tags: Aspose.Cells chart.Title.Font.Name property | Aspose.Cells chart.Title.Font.Size setting | C# bulk chart title formatting Excel | update all chart titles Aspose.Cells | Excel workbook chart title font customization .NET

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an Excel workbook, loops through each worksheet and its charts, ensures each chart title is visible, sets the title font name to Arial and size to 12 points, then saves the modified file.
class SetChartTitlesFont
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
                // Ensure the title is visible (optional)
                chart.Title.IsVisible = true;

                // Set the title font to Arial, size 12
                chart.Title.Font.Name = "Arial";
                chart.Title.Font.Size = 12;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
