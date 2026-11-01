// Title: Add a column chart with Aspose.Cells for .NET, hide its legend, and enable the legend only when a runtime condition is true
// AI Prompts: Generate C# code that uses Aspose.Cells to create a column chart, set ShowLegend = false, and set ShowLegend = true only if a given boolean variable evaluates to true. | Write a method that populates worksheet data, inserts a column chart, hides the legend initially, checks a runtime condition, and toggles the legend visibility accordingly with Aspose.Cells. | Provide an Aspose.Cells example that saves an Excel file where the chart legend appears only when a custom GetShowLegendCondition method returns true.
// Common Searches: aspnet aspose.cells hide chart legend then show based on condition | c# aspose.cells conditional display of chart legend at runtime | how to toggle Excel chart legend visibility using Aspose.Cells .NET
// Tags: Aspose.Cells hide chart legend | Aspose.Cells show chart legend at runtime | Aspose.Cells column chart creation C# | Aspose.Cells chart legend toggle | Excel chart legend control Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// Creates a new workbook, adds month and sales data, inserts a column chart referencing that data, initially hides the chart legend, evaluates a runtime boolean via GetShowLegendCondition, re‑enables the legend if the condition is true, and saves the file as ChartWithConditionalLegend.xlsx.
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

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Month");
            sheet.Cells["B1"].PutValue("Sales");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");
            sheet.Cells["B2"].PutValue(120);
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["B4"].PutValue(130);

            // Add a column chart to the worksheet
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 5);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart series
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Hide the legend initially
            chart.ShowLegend = false;

            // Determine at runtime whether the legend should be visible
            bool showLegend = GetShowLegendCondition();

            // Re‑enable legend visibility if the condition is true
            if (showLegend)
            {
                chart.ShowLegend = true;
            }

            // Save the workbook to a file
            workbook.Save("ChartWithConditionalLegend.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Placeholder for runtime logic that decides legend visibility
    static bool GetShowLegendCondition()
    {
        // Insert actual condition logic here; returning true for demonstration
        return true;
    }
}
