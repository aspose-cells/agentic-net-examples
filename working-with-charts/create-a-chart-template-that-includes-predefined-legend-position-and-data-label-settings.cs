// Title: C# – Create an Excel chart template with predefined legend position and data label formatting using Aspose.Cells
// Description: This example shows how to build a new Workbook, add sample data, insert a column chart, set the legend to the top with custom width, height and bold font, enable value and category name data labels with a dark‑blue font, recalculate the layout, and save the result as an XLSX chart template.
// Keywords: Aspose.Cells chart template C# | set legend position Aspose.Cells | custom data labels Aspose.Cells | Excel chart legend formatting | Aspose.Cells column chart example
// Common Searches: Aspose.Cells how to position chart legend | C# Aspose.Cells data label formatting | create reusable Excel chart template with Aspose.Cells | save chart as template Aspose.Cells .NET | set legend size and font Aspose.Cells
// Developer Intent: Generate an Excel chart template in C# where the legend position, size and style, as well as data‑label visibility and appearance, are predefined.
// Use Cases: Standardize monthly sales charts with a top legend and visible value/category labels. | Provide a reusable chart template that preserves corporate branding for legends and data labels across multiple workbooks. | Automate report generation where every chart must follow specific legend dimensions and label styling.
// AI Prompts: Write C# code with Aspose.Cells to create a line chart template that places the legend at the bottom and shows percentage data labels. | Give a snippet that copies an existing Aspose.Cells chart template, updates its data range, and keeps the original legend and label formatting. | Explain how to set the data label position to OutsideEnd for a bar chart in Aspose.Cells, including any version‑specific considerations.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartTemplate
{
    // This example shows how to build a new Workbook, add sample data, insert a column chart, set the legend to the top with custom width, height and bold font, enable value and category name data labels with a dark‑blue font, recalculate the layout, and save the result as an XLSX chart template.
    public class CreateChartTemplate
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Define the data range for the chart
                chart.SetChartDataRange("A1:B4", true);

                // ---- Predefined Legend Settings ----
                chart.Legend.Position = LegendPositionType.Top;
                chart.Legend.IsOverLay = false;
                chart.Legend.IsAutomaticSize = false;
                chart.Legend.Width = 300;
                chart.Legend.Height = 40;
                chart.Legend.Font.Size = 11;
                chart.Legend.Font.IsBold = true;

                // ---- Predefined Data Label Settings ----
                chart.NSeries[0].DataLabels.ShowValue = true;
                chart.NSeries[0].DataLabels.ShowCategoryName = true;
                // Position data labels outside the data points (if supported by the library version)
                // chart.NSeries[0].DataLabels.Position = DataLabelPositionType.OutsideEnd;
                chart.NSeries[0].DataLabels.Font.Size = 10;
                chart.NSeries[0].DataLabels.Font.Color = System.Drawing.Color.DarkBlue;

                // Apply layout changes
                chart.Calculate();

                // Save the workbook containing the chart template
                workbook.Save("ChartTemplateWithLegendAndDataLabels.xlsx", SaveFormat.Xlsx);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error creating chart template: {ex.Message}");
            }
        }
    }

    public class Program
    {
        public static void Main(string[] args)
        {
            CreateChartTemplate.Run();
        }
    }
}
