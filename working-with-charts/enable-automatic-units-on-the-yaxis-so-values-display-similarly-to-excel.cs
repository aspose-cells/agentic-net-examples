// Title: C# – Enable Automatic Display Units on a Chart Y‑Axis with Aspose.Cells
// Description: A concise Aspose.Cells for .NET example that creates a workbook, adds large numeric data, inserts a column chart, and configures the value (Y) axis to use automatic minimum, maximum, major and minor units while showing the Excel‑style display unit label (e.g., “Millions”). The workbook is saved as an XLSX file.
// Keywords: Aspose.Cells C# chart automatic display unit | Y axis automatic scaling Aspose.Cells | show millions label chart axis | Excel‑like axis units Aspose.Cells | value axis IsDisplayUnitLabelShown | automatic major unit C# | Aspose.Cells chart example GitHub | column chart automatic units
// Common Searches: Aspose.Cells enable automatic display units on Y axis | C# chart axis automatic scaling Aspose.Cells | show millions label on chart using Aspose.Cells | set automatic major unit for value axis C# | Aspose.Cells automatic axis units example
// Developer Intent: Configure a chart’s Y‑axis to automatically choose appropriate units (thousands, millions, etc.) and display the unit label, mimicking Excel behavior.
// Use Cases: Generate financial or sales charts with large numbers where the Y‑axis automatically adjusts to thousands, millions, or billions. | Create reusable reporting templates that adapt axis scaling without manual calculations. | Export dashboards to Excel where the axis label (e.g., “Millions”) appears automatically based on data magnitude.
// AI Prompts: Write C# code using Aspose.Cells to add a line chart and enable automatic display units on its value axis. | Show how to set IsAutomaticMinValue, IsAutomaticMaxValue, and IsDisplayUnitLabelShown for a bar chart in Aspose.Cells. | Explain how to customize the display unit label format after enabling automatic units with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsAutomaticDisplayUnit
{
    // A concise Aspose.Cells for .NET example that creates a workbook, adds large numeric data, inserts a column chart, and configures the value (Y) axis to use automatic minimum, maximum, major and minor units while showing the Excel‑style display unit label (e.g., “Millions”). The workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data (large values to trigger automatic display unit)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");

            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(1500000);   // 1.5 million
            sheet.Cells["B3"].PutValue(3000000);   // 3 million
            sheet.Cells["B4"].PutValue(4500000);   // 4.5 million

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Enable automatic scaling for the value (Y) axis
            Axis valueAxis = chart.ValueAxis;
            valueAxis.IsAutomaticMinValue = true;      // let Excel decide min
            valueAxis.IsAutomaticMaxValue = true;      // let Excel decide max
            valueAxis.IsAutomaticMajorUnit = true;     // let Excel decide major unit
            valueAxis.IsAutomaticMinorUnit = true;     // let Excel decide minor unit

            // Show the display unit label (e.g., "Millions") if Excel determines a unit
            valueAxis.IsDisplayUnitLabelShown = true;

            // Save the workbook
            workbook.Save("AutomaticDisplayUnitChart.xlsx", SaveFormat.Xlsx);
        }
    }
}
