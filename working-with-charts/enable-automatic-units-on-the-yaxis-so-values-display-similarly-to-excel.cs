// Title: Automatic Y‑Axis Display Units in Aspose.Cells Charts (C# .NET)
// Description: Creates a workbook, adds large numbers, inserts a column chart, and configures the ValueAxis to let Excel automatically choose the proper display unit (DisplayUnitType.None) and show the unit label, then saves the file as XLSX.
// Keywords: Aspose.Cells | C# chart display unit | automatic Y‑axis units | ValueAxis DisplayUnitType.None | show axis unit label | Excel‑style scaling | .NET chart formatting
// Common Searches: Aspose.Cells set chart Y axis auto unit | DisplayUnitType.None example C# | show unit label on chart axis Aspose | automatic scaling for large values in Aspose chart | how to enable auto display units Aspose.Cells
// Developer Intent: Configure a chart’s value axis to automatically select and display the appropriate unit label.
// Use Cases: Financial reports with millions‑range figures where the axis should read “Millions” automatically. | Sales dashboards that need readable Y‑axis labels without manually setting a scale. | Reusable chart styling that applies auto‑unit logic to column, line, or bar charts across projects.
// AI Prompts: Generate C# code that creates a line chart in Aspose.Cells with automatic Y‑axis display units and a visible unit label. | Explain the effect of DisplayUnitType.None on a chart axis and how to toggle the unit‑label visibility in Aspose.Cells. | Provide a helper method for any Aspose.Cells chart that enables automatic display units on its value axis.

using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, adds large numbers, inserts a column chart, and configures the ValueAxis to let Excel automatically choose the proper display unit (DisplayUnitType.None) and show the unit label, then saves the file as XLSX.
class AutomaticDisplayUnitDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add sample data with large values
        worksheet.Cells["A1"].PutValue("Category");
        worksheet.Cells["A2"].PutValue("A");
        worksheet.Cells["A3"].PutValue("B");
        worksheet.Cells["A4"].PutValue("C");
        worksheet.Cells["B1"].PutValue("Value");
        worksheet.Cells["B2"].PutValue(1500000);
        worksheet.Cells["B3"].PutValue(3000000);
        worksheet.Cells["B4"].PutValue(4500000);

        // Insert a column chart
        int chartIndex = worksheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
        Chart chart = worksheet.Charts[chartIndex];

        // Set the data range for the chart
        chart.NSeries.Add("B2:B4", true);
        chart.NSeries.CategoryData = "A2:A4";

        // Enable automatic display units on the Y‑axis
        // Setting DisplayUnit to None lets Excel decide the appropriate unit.
        // Showing the unit label makes the chosen unit visible on the axis.
        chart.ValueAxis.DisplayUnit = DisplayUnitType.None;
        chart.ValueAxis.IsDisplayUnitLabelShown = true;

        // Save the workbook
        workbook.Save("AutomaticDisplayUnitDemo.xlsx", SaveFormat.Xlsx);
    }
}
