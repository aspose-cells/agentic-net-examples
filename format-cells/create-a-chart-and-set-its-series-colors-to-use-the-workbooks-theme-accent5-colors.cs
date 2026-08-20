// Title: Aspose.Cells for .NET: Apply Workbook Theme Accent5 to All Chart Series (C#)
// Description: Creates a workbook, adds sample data, inserts a column chart, and assigns the workbook’s Accent5 theme to the fill and border of every series before saving the file.
// Keywords: Aspose.Cells | C# chart theme color | Accent5 | ThemeColorType.Accent5 | set chart series color | Aspose.Cells chart styling | Excel chart theme color .NET | solid fill chart series | Aspose.Cells ThemeColor | column chart series color
// Common Searches: Aspose.Cells set chart series Accent5 | C# apply workbook theme color to chart series | How to use ThemeColor Accent5 in Aspose.Cells | Change chart series fill to theme accent in .NET | Aspose.Cells chart series color example
// Developer Intent: Use the workbook’s Accent5 theme to color all series of a chart.
// Use Cases: Generate a column chart from worksheet data and automatically style each series with the workbook’s Accent5 color. | Maintain consistent branding by applying the same theme accent to series fills and borders. | Create Excel reports where chart colors adapt to the workbook’s theme without hard‑coding RGB values. | Programmatically update existing charts to match a new theme across multiple series.
// AI Prompts: Show C# code that sets a chart series’ fill and border to the workbook’s Accent5 ThemeColor using Aspose.Cells. | Explain how to retrieve and assign ThemeColorType.Accent5 to series.Area.FillFormat.SolidFill.CellsColor. | Provide a step‑by‑step guide to iterate over chart.NSeries and apply a solid fill with Accent5. | What methods are needed to change a chart series border to the workbook’s Accent5 theme color in Aspose.Cells?

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;
using System.Drawing;

namespace AsposeCellsChartAccent5Demo
{
    // Creates a workbook, adds sample data, inserts a column chart, and assigns the workbook’s Accent5 theme to the fill and border of every series before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["A2"].PutValue("Jan");
            ws.Cells["A3"].PutValue("Feb");
            ws.Cells["A4"].PutValue("Mar");
            ws.Cells["A5"].PutValue("Apr");

            ws.Cells["B1"].PutValue("Series1");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);
            ws.Cells["B5"].PutValue(40);

            ws.Cells["C1"].PutValue("Series2");
            ws.Cells["C2"].PutValue(15);
            ws.Cells["C3"].PutValue(25);
            ws.Cells["C4"].PutValue(35);
            ws.Cells["C5"].PutValue(45);

            // Add a column chart
            int chartIdx = ws.Charts.Add(ChartType.Column, 6, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series (both series)
            chart.NSeries.Add("B2:C5", true);
            chart.NSeries.CategoryData = "A2:A5";

            // Apply the workbook's Accent5 theme color to each series
            for (int i = 0; i < chart.NSeries.Count; i++)
            {
                Series series = chart.NSeries[i];

                // Ensure the fill type is solid
                series.Area.FillFormat.FillType = FillType.Solid;

                // Set the solid fill to use Theme Accent5
                var cellsColor = series.Area.FillFormat.SolidFill.CellsColor;
                cellsColor.ThemeColor = new ThemeColor(ThemeColorType.Accent5, 0);
                series.Area.FillFormat.SolidFill.CellsColor = cellsColor;

                // Optionally, also set the border to use the same theme color
                series.Border.ThemeColor = new ThemeColor(ThemeColorType.Accent5, 0);
            }

            // Save the workbook
            workbook.Save("ChartAccent5Demo.xlsx");
        }
    }
}
