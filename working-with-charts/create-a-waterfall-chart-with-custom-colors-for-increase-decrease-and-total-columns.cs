// Title: Create a Waterfall Chart with Custom Increase, Decrease, and Total Colors using Aspose.Cells for .NET (C#)
// Description: This example shows how to generate a workbook, populate category and value columns, add a Waterfall chart, and programmatically set point colors—green for positive changes, red for negative changes, and blue for the start/end totals—then save the file as an XLSX document.
// Keywords: Aspose.Cells waterfall chart C# | custom colors waterfall chart Aspose | set chart point color Aspose.Cells | increase decrease total colors | waterfall chart formatting .NET | Aspose.Cells chart point ForegroundColor | C# generate waterfall chart
// Common Searches: how to change column colors in Aspose.Cells waterfall chart | Aspose.Cells set green for positive values | Aspose.Cells red for negative waterfall points | apply blue to total columns Aspose.Cells | C# waterfall chart custom colors example
// Developer Intent: Create a waterfall chart and assign distinct colors to increase, decrease, and total columns programmatically.
// Use Cases: Financial statements: highlight revenue gains (green), losses (red), and opening/closing balances (blue). | Project budgeting: visualize each cost impact with color‑coded columns for quick analysis. | Automated reporting: generate Excel reports where waterfall chart colors adapt to underlying data values.
// AI Prompts: Generate C# code with Aspose.Cells that builds a waterfall chart and colors points based on positive, negative, and total values. | Explain step‑by‑step how to access a chart series in Aspose.Cells and set the ForegroundColor for each ChartPoint. | Show how to detect the first and last points of a waterfall series and apply a unique color using Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // This example shows how to generate a workbook, populate category and value columns, add a Waterfall chart, and programmatically set point colors—green for positive changes, red for negative changes, and blue for the start/end totals—then save the file as an XLSX document.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate data for a waterfall chart
                // Column A – Categories, Column B – Values
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Start");
                sheet.Cells["A3"].PutValue("Increase 1");
                sheet.Cells["A4"].PutValue("Decrease 1");
                sheet.Cells["A5"].PutValue("Increase 2");
                sheet.Cells["A6"].PutValue("Total");

                // Values: first and last rows are totals, others are changes
                sheet.Cells["B2"].PutValue(100);   // Start total
                sheet.Cells["B3"].PutValue(30);    // Increase
                sheet.Cells["B4"].PutValue(-20);   // Decrease
                sheet.Cells["B5"].PutValue(50);    // Increase
                sheet.Cells["B6"].PutValue(160);   // Final total (calculated manually)

                // Add a Waterfall chart
                int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 8, 0, 25, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series (values only)
                chart.NSeries.Add("B2:B6", true);
                // Set the category (X‑axis) data
                chart.NSeries.CategoryData = "A2:A6";

                // Customize colors for each point
                // Increase – Green, Decrease – Red, Total – Blue
                Series series = chart.NSeries[0];
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // For total columns (first and last points) use blue
                    if (i == 0 || i == series.Points.Count - 1)
                    {
                        point.Area.ForegroundColor = Color.Blue;
                    }
                    else
                    {
                        // Retrieve the underlying cell value to decide color
                        // Column B (index 1), rows start at 2 (index 1)
                        double val = Convert.ToDouble(sheet.Cells[i + 1, 1].Value);
                        point.Area.ForegroundColor = val >= 0 ? Color.Green : Color.Red;
                    }
                }

                // Optional: add a title
                chart.Title.Text = "Waterfall Chart with Custom Colors";

                // Save the workbook
                string outputPath = "WaterfallCustomColors.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
