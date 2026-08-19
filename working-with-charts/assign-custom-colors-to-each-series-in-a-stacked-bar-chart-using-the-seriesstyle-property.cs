// Title: Assign Custom Colors to Series in a Stacked Bar Chart – Aspose.Cells for .NET (C#)
// Description: Creates a workbook, fills it with sample data, adds a stacked bar chart, and applies a unique color to each series using the Series.Style (Area.ForegroundColor) property before saving the file as an Excel workbook.
// Keywords: Aspose.Cells | C# | stacked bar chart | custom series colors | Series.Style | Area.ForegroundColor | Excel chart formatting | chart series color .NET | Aspose.Cells example
// Common Searches: Aspose.Cells set individual series color stacked bar | C# change series color in Aspose.Cells chart | how to use Series.Style for chart colors Aspose.Cells | apply custom colors to stacked bar chart series .NET | Aspose.Cells chart series formatting tutorial
// Developer Intent: Apply a distinct custom color to each series of a stacked bar chart in an Excel file using Aspose.Cells for .NET.
// Use Cases: Financial reports where each product line uses brand‑specific colors in a stacked bar chart. | Sales dashboards that match corporate palettes for clear visual separation of series. | Performance presentations that require differentiated colors for multi‑category stacked bars.
// AI Prompts: Show a C# snippet that assigns custom colors to each series of a stacked bar chart with Aspose.Cells and saves the workbook. | Explain how to map a Color array to chart series using Series.Style or Area.ForegroundColor in Aspose.Cells. | Provide step‑by‑step instructions for coloring individual series in a stacked bar chart with Aspose.Cells for .NET.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates a workbook, fills it with sample data, adds a stacked bar chart, and applies a unique color to each series using the Series.Style (Area.ForegroundColor) property before saving the file as an Excel workbook.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for a stacked bar chart
        sheet.Cells["A1"].PutValue("Category");
        sheet.Cells["A2"].PutValue("Q1");
        sheet.Cells["A3"].PutValue("Q2");
        sheet.Cells["A4"].PutValue("Q3");

        sheet.Cells["B1"].PutValue("Series1");
        sheet.Cells["B2"].PutValue(10);
        sheet.Cells["B3"].PutValue(20);
        sheet.Cells["B4"].PutValue(30);

        sheet.Cells["C1"].PutValue("Series2");
        sheet.Cells["C2"].PutValue(15);
        sheet.Cells["C3"].PutValue(25);
        sheet.Cells["C4"].PutValue(35);

        sheet.Cells["D1"].PutValue("Series3");
        sheet.Cells["D2"].PutValue(12);
        sheet.Cells["D3"].PutValue(22);
        sheet.Cells["D4"].PutValue(32);

        // Add a stacked bar chart to the worksheet
        int chartIdx = sheet.Charts.Add(ChartType.BarStacked, 5, 0, 20, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Set the data range for the series and the categories
        chart.NSeries.Add("B2:D4", true);          // Series values
        chart.NSeries.CategoryData = "A2:A4";      // Category (X‑axis) values

        // Define custom colors for each series
        Color[] customColors = new Color[]
        {
            Color.FromArgb(79, 129, 189),   // First series color
            Color.FromArgb(192, 80, 77),   // Second series color
            Color.FromArgb(155, 187, 89)   // Third series color
        };

        // Apply the custom colors to each series using the Area.ForegroundColor property
        for (int i = 0; i < chart.NSeries.Count; i++)
        {
            chart.NSeries[i].Area.ForegroundColor = customColors[i];
        }

        // Save the workbook with the customized stacked bar chart
        workbook.Save("StackedBarCustomColors.xlsx");
    }
}
