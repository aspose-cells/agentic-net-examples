// Title: Aspose.Cells C# – Set Primary Y‑Axis Major Unit to 10 for a Column Chart
// Description: Creates a new workbook, populates A1:B6 with categories and values, adds a column chart, binds the data range, disables automatic major‑unit calculation on the ValueAxis, sets the major unit to 10, and saves the file as ChartWithMajorUnit.xlsx.
// Keywords: Aspose.Cells set major unit | C# chart Y axis interval | primary value axis grid spacing | Aspose.Cells chart scaling | column chart major unit 10 | .NET Excel chart customization
// Common Searches: Aspose.Cells fix Y axis major unit C# | set fixed major grid interval for Excel chart using Aspose | disable automatic major unit Aspose.Cells .NET | how to standardize chart axis intervals in C# | chart ValueAxis MajorUnit example Aspose
// Developer Intent: Configure a column chart so the primary Y‑axis uses a fixed major unit of 10 instead of the automatic calculation.
// Use Cases: Align grid lines across multiple charts in a reporting template. | Match axis intervals to business thresholds such as multiples of 10. | Create reproducible chart layouts for downstream Excel analysis.
// AI Prompts: Write C# code with Aspose.Cells to generate a line chart and set its primary Y‑axis major unit to 5. | Explain how to turn off IsAutomaticMajorUnit and assign a custom MajorUnit for any chart type in Aspose.Cells .NET. | Provide a script that iterates through all worksheets in a workbook and applies a major unit of 10 to every chart's value axis.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a new workbook, populates A1:B6 with categories and values, adds a column chart, binds the data range, disables automatic major‑unit calculation on the ValueAxis, sets the major unit to 10, and saves the file as ChartWithMajorUnit.xlsx.
class SetMajorUnitExample
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        for (int i = 1; i <= 5; i++)
        {
            sheet.Cells[i + 1, 0].PutValue("Item " + i);   // Category labels
            sheet.Cells[i + 1, 1].PutValue(i * 12);       // Sample numeric values
        }

        // Add a column chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Bind the data range to the chart
        chart.NSeries.Add("B2:B6", true);          // Values
        chart.NSeries.CategoryData = "A2:A6";     // Categories

        // Set the primary Y‑axis (ValueAxis) major unit to 10
        chart.ValueAxis.IsAutomaticMajorUnit = false; // Disable auto calculation
        chart.ValueAxis.MajorUnit = 10;                // Standardize grid interval

        // Save the workbook with the configured chart
        workbook.Save("ChartWithMajorUnit.xlsx");
    }
}
