// Title: Apply Built‑In Chart Style Style1 to All Charts in an Aspose.Cells Workbook (C#)
// Description: This example creates a workbook, adds sample data and two charts (column and pie), then loops through every worksheet and each chart to set the chart's Style property to the built‑in Style1 (value = 1). The workbook is saved as an XLSX file with a consistent visual theme across all charts.
// Keywords: Aspose.Cells chart style C# | Style1 chart Aspose | apply built‑in chart style | set chart style programmatically | iterate worksheets charts Aspose.Cells | C# Aspose.Cells visual formatting | chart style enumeration Aspose | global chart theme workbook
// Common Searches: how to set Style1 for every chart in Aspose.Cells C# | apply predefined chart style to all charts Aspose.Cells | change chart appearance programmatically Aspose.Cells .NET | loop through worksheets and charts to set style | Aspose.Cells chart style property example
// Developer Intent: Set the built‑in Style1 on every chart in a workbook using Aspose.Cells for C#.
// Use Cases: Standardize chart appearance across a multi‑sheet report before distribution. | Enforce corporate branding by applying a single chart style to all generated workbooks. | Update legacy workbooks that contain newly added charts so they automatically adopt Style1.
// AI Prompts: Generate C# code with Aspose.Cells to apply Style2 to all charts in a workbook. | Show how to list all available chart style IDs and let the user choose one to apply to specific chart types. | Provide an example of loading a custom chart style from a template and applying it to charts across multiple worksheets.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartStyleDemo
{
    // This example creates a workbook, adds sample data and two charts (column and pie), then loops through every worksheet and each chart to set the chart's Style property to the built‑in Style1 (value = 1). The workbook is saved as an XLSX file with a consistent visual theme across all charts.
    class Program
    {
        static void Main()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add sample data for charts
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["B3"].PutValue(20);
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["B4"].PutValue(30);

            // Add a few charts to demonstrate the style application
            int chartIdx1 = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
            Chart chart1 = sheet.Charts[chartIdx1];
            chart1.NSeries.Add("B2:B4", false);
            chart1.NSeries.CategoryData = "A2:A4";

            int chartIdx2 = sheet.Charts.Add(ChartType.Pie, 5, 10, 20, 18);
            Chart chart2 = sheet.Charts[chartIdx2];
            chart2.NSeries.Add("B2:B4", false);
            chart2.NSeries.CategoryData = "A2:A4";

            // Apply the predefined builtin style 'Style1' (value = 1) to all charts in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                foreach (Chart ch in ws.Charts)
                {
                    ch.Style = 1; // Style1 corresponds to integer value 1
                }
            }

            // Save the workbook
            workbook.Save("WorkbookWithStyledCharts.xlsx", SaveFormat.Xlsx);
        }
    }
}
