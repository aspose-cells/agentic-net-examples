// Title: C# – Set Primary Y‑Axis Major Unit to 10 in an AspNet Aspose.Cells Column Chart
// Description: Creates a workbook, adds sample data, inserts a column chart, disables automatic major‑unit calculation on the primary Y axis, sets the major unit to 10, and saves the file as ChartWithStandardizedYAxis.xlsx.
// Keywords: Aspose.Cells Y axis major unit | C# chart axis scaling | set primary Y axis interval | disable automatic major unit Aspose | column chart grid lines | .NET Excel chart formatting
// Common Searches: Aspose.Cells set Y axis major unit C# | how to fix primary Y axis interval in Aspose chart | disable automatic major unit Aspose.Cells .NET | standardize chart grid lines Aspose.Cells | C# example for manual Y axis scaling
// Developer Intent: Manually define the primary Y‑axis major unit as 10 to keep chart grid lines consistent.
// Use Cases: Building dashboards where all charts share the same Y‑axis scale for easy comparison. | Generating reports that require Y‑axis ticks aligned with business thresholds (e.g., multiples of 10). | Automating Excel exports where precise axis intervals improve readability across multiple workbooks.
// AI Prompts: Provide C# code to set a fixed major unit of 10 on the primary Y axis of an Aspose.Cells chart. | Show how to turn off automatic major‑unit calculation for a chart's ValueAxis in Aspose.Cells for .NET. | Explain how to adjust Y‑axis scaling for existing charts in a workbook using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds sample data, inserts a column chart, disables automatic major‑unit calculation on the primary Y axis, sets the major unit to 10, and saves the file as ChartWithStandardizedYAxis.xlsx.
class SetPrimaryYAxisMajorUnit
{
    static void Main()
    {
        // Create a new workbook (lifecycle: create)
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for the chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["B1"].PutValue("Value");
        sheet.Cells["A2"].PutValue("A");
        sheet.Cells["B2"].PutValue(12);
        sheet.Cells["A3"].PutValue("B");
        sheet.Cells["B3"].PutValue(27);
        sheet.Cells["A4"].PutValue("C");
        sheet.Cells["B4"].PutValue(45);
        sheet.Cells["A5"].PutValue("D");
        sheet.Cells["B5"].PutValue(63);

        // Add a column chart
        int chartIndex = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 12);
        Chart chart = sheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B5", true);
        chart.NSeries.CategoryData = "A2:A5";

        // Disable automatic major unit calculation and set the major unit to 10
        chart.ValueAxis.IsAutomaticMajorUnit = false;   // ensure manual setting is respected
        chart.ValueAxis.MajorUnit = 10;                // major unit interval on primary Y axis

        // Save the workbook (lifecycle: save)
        workbook.Save("ChartWithStandardizedYAxis.xlsx");
    }
}
