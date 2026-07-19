// Title: Aspose.Cells for .NET – Set Custom RGB Colors for Chart Series (C#)
// Description: Creates a workbook, adds a column chart with two series, defines an array of Color objects using specific RGB values, and applies each color to the corresponding series via Area.ForegroundColor and a solid FormattingType.Custom fill before saving the file.
// Keywords: Aspose.Cells C# chart series color | custom RGB palette Aspose.Cells | Area.ForegroundColor .NET | FormattingType.Custom chart fill | column chart series styling | Excel chart color programmatically | Aspose.Cells chart customization | set series fill color .NET
// Common Searches: Aspose.Cells set series RGB color C# | how to change chart series fill Aspose.Cells | custom color palette for Excel chart using Aspose | apply solid fill to chart series Aspose.Cells .NET | programmatic chart styling Aspose.Cells
// Developer Intent: Apply specific RGB colors to each series of an Excel chart using Aspose.Cells.
// Use Cases: Brand‑compliant reporting: assign corporate palette colors to chart series. | Financial dashboards: differentiate revenue and expense series with distinct hues. | Automated workbook generation: ensure consistent visual styling across multiple chart types.
// AI Prompts: Show C# code to set custom RGB colors for each series in an Aspose.Cells chart. | Explain how to use Area.ForegroundColor and FormattingType.Custom to style chart series. | Provide a step‑by‑step guide for applying a custom color palette to column and line charts with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace CustomChartPaletteDemo
{
    // Creates a workbook, adds a column chart with two series, defines an array of Color objects using specific RGB values, and applies each color to the corresponding series via Area.ForegroundColor and a solid FormattingType.Custom fill before saving the file.
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for two series
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["A2"].PutValue("Jan");
            sheet.Cells["A3"].PutValue("Feb");
            sheet.Cells["A4"].PutValue("Mar");

            sheet.Cells["B1"].PutValue("Series1");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["B4"].PutValue(30);

            sheet.Cells["C1"].PutValue("Series2");
            sheet.Cells["C2"].PutValue(15);
            sheet.Cells["C3"].PutValue(25);
            sheet.Cells["C4"].PutValue(35);

            // Add a column chart
            int chartIdx = sheet.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = sheet.Charts[chartIdx];

            // Set the data range for the series (two series)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom RGB colors for each series
            Color[] customSeriesColors = new Color[]
            {
                Color.FromArgb(79, 129, 189),   // Color for Series1
                Color.FromArgb(192, 80, 77)     // Color for Series2
            };

            // Apply the custom colors to each series using the Area.ForegroundColor property
            for (int i = 0; i < chart.NSeries.Count && i < customSeriesColors.Length; i++)
            {
                // Set the fill color of the series area (affects column fill)
                chart.NSeries[i].Area.ForegroundColor = customSeriesColors[i];
                // Ensure the area uses a solid fill
                chart.NSeries[i].Area.Formatting = FormattingType.Custom;
            }

            // Save the workbook (using the provided lifecycle rule)
            workbook.Save("CustomChartPaletteDemo.xlsx");
        }
    }
}
