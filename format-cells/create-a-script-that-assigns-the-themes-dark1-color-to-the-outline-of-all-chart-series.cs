// Title: C# – Apply Theme Dark1 Color to Chart Series Outlines with Aspose.Cells
// Description: Demonstrates how to create a workbook, add sample data, insert a column chart, and set the outline of every chart series to the workbook's Dark1 (Background1) theme color using Aspose.Cells for .NET. The example also shows how to make the border visible and assign a medium line weight before saving the file.
// Keywords: Aspose.Cells chart series border | C# set chart outline theme color | Dark1 theme color Aspose.Cells | ThemeColor Background1 chart series | .NET chart formatting Aspose | apply theme to chart series outline | Aspose.Cells column chart styling
// Common Searches: Aspose.Cells set chart series outline Dark1 | C# Aspose.Cells theme color for chart borders | How to use ThemeColor for chart series in Aspose.Cells | Apply workbook theme to chart series outline .NET | Change chart series border to Background1 color
// Developer Intent: Programmatically set the outline of all chart series to the workbook's Dark1 (Background1) theme color using Aspose.Cells for .NET.
// Use Cases: Generate reports where chart series outlines follow the document theme for brand consistency. | Create column charts with visible, theme‑aligned borders for better visual hierarchy. | Update existing workbooks to ensure chart series borders match the selected theme before distribution.
// AI Prompts: Write C# code with Aspose.Cells that assigns the Dark1 (Background1) theme color to every chart series outline and makes the border visible. | Show how to set a medium line weight for chart series borders after applying the Dark1 theme color in Aspose.Cells. | Explain how to retrieve other theme colors such as Accent1 or Dark2 and apply them to chart series borders using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsThemeOutlineDemo
{
    // Demonstrates how to create a workbook, add sample data, insert a column chart, and set the outline of every chart series to the workbook's Dark1 (Background1) theme color using Aspose.Cells for .NET. The example also shows how to make the border visible and assign a medium line weight before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Series 1");
            sheet.Cells["C1"].PutValue("Series 2");

            for (int i = 2; i <= 6; i++)
            {
                sheet.Cells[$"A{i}"].PutValue($"Cat {i - 1}");
                sheet.Cells[$"B{i}"].PutValue(i * 10);
                sheet.Cells[$"C{i}"].PutValue(i * 15);
            }

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data range for the series and categories
            chart.NSeries.Add("B1:C6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Assign the theme's Dark1 color (Background1) to the outline of each series
            ThemeColor dark1Theme = new ThemeColor(ThemeColorType.Background1, 0.0);
            foreach (Series series in chart.NSeries)
            {
                series.Border.ThemeColor = dark1Theme;   // Set theme color for the border
                series.Border.IsVisible = true;          // Ensure the border is drawn
                series.Border.Weight = WeightType.MediumLine; // Optional: set line weight
            }

            // Save the workbook
            workbook.Save("ChartSeriesOutlineWithDark1Theme.xlsx");
        }
    }
}
