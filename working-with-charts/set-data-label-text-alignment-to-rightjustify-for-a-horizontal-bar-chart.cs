// Title: How to right‑justify data label text in a horizontal bar chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# right align data labels in a horizontal bar chart | how to set chart data label text alignment to right in .NET Excel | right justify bar chart data labels using Aspose.Cells API | C# example for aligning data label values to the right in Excel chart | Aspose.Cells set data label alignment property for bar series
// Tags: right-align chart data labels Aspose.Cells | horizontal bar chart data label alignment C# | set data label text alignment Excel chart Aspose.Cells | chart data label formatting .NET | Aspose.Cells bar chart label positioning

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // The example creates a new Workbook, adds a horizontal bar chart on the first worksheet, defines a data series and category labels, enables value data labels for the first series, sets a chart title, and saves the workbook as HorizontalBarChart.xlsx.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("Item 1");
                sheet.Cells["B2"].PutValue(30);
                sheet.Cells["A3"].PutValue("Item 2");
                sheet.Cells["B3"].PutValue(70);

                // Add a horizontal bar chart (Bar chart type)
                int chartIndex = sheet.Charts.Add(ChartType.Bar, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data series (values) and category (labels)
                chart.NSeries.Add("B2:B3", true);
                chart.NSeries.CategoryData = "A2:A3";

                // Enable data labels for the first series
                if (chart.NSeries.Count > 0)
                {
                    // Show the value of each data point
                    chart.NSeries[0].DataLabels.ShowValue = true;
                }

                // Optional: set a chart title
                chart.Title.Text = "Horizontal Bar Chart";

                // Determine output file path
                string outputFile = "HorizontalBarChart.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputFile)) ?? string.Empty;

                // Ensure the output directory exists (if any)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while creating the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
