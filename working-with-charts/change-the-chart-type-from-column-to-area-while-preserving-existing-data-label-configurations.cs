// Title: Convert a Column Chart to an Area Chart while Keeping Data Label Settings – Aspose.Cells for .NET
// Description: Shows how to build a workbook, add a column chart with custom data labels (value, category name, shape, background color), change the chart type to Area via Chart.Type, and retain all label formatting before saving the file as XLSX.
// Keywords: Aspose.Cells | C# | .NET | chart type conversion | column to area chart | data label formatting | preserve data labels | Chart.Type property | Excel automation | programmatic chart modification
// Common Searches: Aspose.Cells change column chart to area chart | keep data label formatting when switching chart type | C# convert Excel column chart to area chart | preserve chart data labels Aspose.Cells | how to modify chart type programmatically .NET
// Developer Intent: Replace an existing column chart with an area chart while retaining every data‑label property that was previously set.
// Use Cases: Create a column chart, style its data labels, then switch to an area chart for a different visual presentation without re‑applying label settings. | Batch‑process workbooks where users can select chart styles; the code updates chart types on the fly while preserving label appearance. | Automate a reporting pipeline that dynamically changes chart types based on data trends, ensuring consistent label formatting across all outputs.
// AI Prompts: Generate C# code using Aspose.Cells that converts a column chart to an area chart and keeps all data label options (ShowValue, ShowCategoryName, ShapeType, background color). | Explain why Aspose.Cells retains data label configurations when the Chart.Type property is changed and what limitations, if any, exist.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsChartTypeChange
{
    // Shows how to build a workbook, add a column chart with custom data labels (value, category name, shape, background color), change the chart type to Area via Chart.Type, and retain all label formatting before saving the file as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            // Add a Column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels (these settings should be retained after type change)
            Series series = chart.NSeries[0];
            series.DataLabels.ShowValue = true;                     // Show the value on each data point
            series.DataLabels.ShowCategoryName = true;              // Show category name
            series.DataLabels.ShapeType = DataLabelShapeType.Rect;  // Use rectangular data labels
            series.DataLabels.Area.ForegroundColor = Color.LightGreen; // Example formatting

            // Change the chart type from Column to Area while preserving data label settings
            chart.Type = ChartType.Area;

            // Save the workbook
            workbook.Save("ChartColumnToArea.xlsx", SaveFormat.Xlsx);
        }
    }
}
