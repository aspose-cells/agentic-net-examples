// Title: Programmatically set a custom color palette and apply distinct colors to each chart series in Aspose.Cells for .NET
// AI Prompts: Create a column chart, modify the workbook palette with custom RGB values, and assign each series' Area.ForegroundColor to the matching palette entry using Aspose.Cells. | Change specific palette indices in an Aspose.Cells workbook and bind those colors to individual chart series for solid fills. | Set the Area.Formatting property to Custom for each series after applying palette colors to ensure a solid fill. | Save the workbook containing the customized chart colors to an .xlsx file with Aspose.Cells.
// Common Searches: Aspose.Cells .NET how to use a custom color palette for chart series colors | set different colors for each series in an Excel chart using Aspose.Cells C# | change Excel palette indices programmatically with Aspose.Cells and apply to a chart | assign solid fill colors to chart series from a custom palette in Aspose.Cells | C# example of customizing chart series colors via workbook palette in Aspose.Cells
// Tags: custom workbook palette Aspose.Cells | chart series color assignment Aspose.Cells .NET | column chart custom series fill | change Excel palette index programmatically | Area.ForegroundColor Aspose.Cells | solid fill formatting chart series

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Drawing;

namespace AsposeCellsCustomPaletteDemo
{
    // The sample creates a workbook with sample data, adds a column chart, defines two custom colors in the workbook palette, and assigns those palette colors to the foreground of each chart series. It also sets each series to use a solid fill before saving the file as CustomPaletteChart.xlsx.
    class Program
    {
        static void Main()
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

            // Set the data range for the series (both series in one call)
            chart.NSeries.Add("B1:C4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Define custom colors in the workbook palette (indices 0 and 1)
            // These colors will be used later for the series
            workbook.ChangePalette(Color.FromArgb(79, 129, 189), 0); // Custom blue
            workbook.ChangePalette(Color.FromArgb(192, 80, 77), 1); // Custom red

            // Assign specific colors to each series using the Area.ForegroundColor property
            // Series 0 will use palette index 0, Series 1 will use palette index 1
            chart.NSeries[0].Area.ForegroundColor = workbook.Colors[0];
            chart.NSeries[1].Area.ForegroundColor = workbook.Colors[1];

            // Optionally, ensure that each series uses its own solid fill
            chart.NSeries[0].Area.Formatting = FormattingType.Custom;
            chart.NSeries[1].Area.Formatting = FormattingType.Custom;

            // Save the workbook
            workbook.Save("CustomPaletteChart.xlsx");
        }
    }
}
