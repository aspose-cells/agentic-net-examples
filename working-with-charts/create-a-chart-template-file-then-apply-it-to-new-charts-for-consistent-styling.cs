// Title: Create a reusable chart template and apply consistent styling to multiple charts using Aspose.Cells for .NET (C#)
// Common Searches: how to create a chart template in Aspose.Cells C# and reuse it for other charts | apply the same title font and series color to multiple Aspose.Cells charts programmatically | copy chart formatting from one Aspose.Cells chart to another in a .NET workbook | set data labels to show values on all column charts using Aspose.Cells C# | save an Excel workbook with two identically styled column charts using Aspose.Cells
// Tags: Aspose.Cells chart template reuse C# | Aspose.Cells copy chart styling programmatically | Aspose.Cells set chart title font color | Aspose.Cells add data labels to series | Aspose.Cells create column chart from range | Aspose.Cells save workbook with multiple styled charts

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsChartTemplateDemo
{
    // The example creates a workbook, populates sample data, adds a column chart with custom title, font color, size, series color, and data labels, then creates a second chart and copies the same styling properties from the first chart to achieve consistent appearance, finally saving the workbook as an XLSX file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // -----------------------------------------------------------------
                // 1. Create a workbook and add sample data for the chart
                // -----------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate some data (A1:A5 and B1:B5)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");
                sheet.Cells["A5"].PutValue("Apr");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);
                sheet.Cells["B5"].PutValue(40);

                // -----------------------------------------------------------------
                // 2. Create a chart and apply custom styling
                // -----------------------------------------------------------------
                int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 20, 10);
                Chart chart = sheet.Charts[chartIndex];

                // Set chart data range
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // Custom styling (example)
                chart.Title.Text = "Sales by Month";
                chart.Title.Font.Color = Color.DarkBlue;
                chart.Title.Font.Size = 14;

                // Set the first series color
                chart.NSeries[0].Area.ForegroundColor = Color.CornflowerBlue;

                // Add data labels (show values)
                chart.NSeries[0].DataLabels.ShowValue = true;

                // -----------------------------------------------------------------
                // 3. Create a second chart and apply the same styling manually
                // -----------------------------------------------------------------
                int newChartIndex = sheet.Charts.Add(ChartType.Column, 22, 0, 35, 10);
                Chart newChart = sheet.Charts[newChartIndex];

                // Set data for the new chart (reuse same data range for demo)
                newChart.NSeries.Add("B2:B5", true);
                newChart.NSeries.CategoryData = "A2:A5";

                // Apply the same styling as the first chart
                newChart.Title.Text = chart.Title.Text;
                newChart.Title.Font.Color = chart.Title.Font.Color;
                newChart.Title.Font.Size = chart.Title.Font.Size;
                newChart.NSeries[0].Area.ForegroundColor = chart.NSeries[0].Area.ForegroundColor;
                newChart.NSeries[0].DataLabels.ShowValue = true;

                // -----------------------------------------------------------------
                // 4. Save the workbook to verify both charts are present
                // -----------------------------------------------------------------
                string workbookPath = "ChartTemplateDemo.xlsx";

                // Ensure the directory exists before saving
                string workbookDir = Path.GetDirectoryName(Path.GetFullPath(workbookPath));
                if (!string.IsNullOrEmpty(workbookDir) && !Directory.Exists(workbookDir))
                {
                    Directory.CreateDirectory(workbookDir);
                }

                try
                {
                    workbook.Save(workbookPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Workbook saved successfully to '{workbookPath}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
