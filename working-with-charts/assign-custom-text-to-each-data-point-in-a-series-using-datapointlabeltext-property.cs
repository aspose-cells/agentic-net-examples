// Title: C# – Set custom text for each chart data point using DataLabels.Text in Aspose.Cells
// Description: Creates a workbook, adds a column chart from cells A2:A4 and B2:B4, enables data labels, disables automatic text, and assigns a custom label that combines the category name from column A with the point's Y‑value via the DataLabels.Text property. The workbook is saved as CustomDataPointLabels.xlsx.
// Keywords: Aspose.Cells custom data point labels | DataLabels.Text C# example | chart point label Aspose.Cells | column chart custom labels .NET | set chart point text programmatically
// Common Searches: Aspose.Cells set custom label for each chart point | C# example DataLabels.Text chart | how to display category name with value in Aspose chart | custom data labels Aspose.Cells .NET | retrieve cell value for chart point label
// Developer Intent: Assign a unique text label to every data point in a chart series using Aspose.Cells for .NET.
// Use Cases: Show product codes and sales figures together on each column bar for a sales dashboard. | Display department names alongside performance metrics in a KPI chart. | Combine row identifiers with measurement values for detailed scientific reporting.
// AI Prompts: Generate code to change the column chart to a line chart while preserving custom data point labels. | Show how to apply individual font colors and sizes to each custom data label in Aspose.Cells. | Explain how to pull custom label text from a separate worksheet range instead of building the string in code.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart from cells A2:A4 and B2:B4, enables data labels, disables automatic text, and assigns a custom label that combines the category name from column A with the point's Y‑value via the DataLabels.Text property. The workbook is saved as CustomDataPointLabels.xlsx.
    public class CustomDataPointLabelsDemo
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
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 8);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for each series and assign custom text
                foreach (Series series in chart.NSeries)
                {
                    series.DataLabels.ShowValue = true; // show the default value

                    for (int i = 0; i < series.Points.Count; i++)
                    {
                        ChartPoint point = series.Points[i];
                        point.DataLabels.IsAutoText = false;

                        // Retrieve the category name from the worksheet (A2:A4)
                        string category = sheet.Cells[i + 1, 0].StringValue; // row i+1, column 0 (A)

                        // Set custom label text
                        point.DataLabels.Text = $"Item {category}: {point.YValue}";
                    }
                }

                // Define output file path
                string outputPath = Path.Combine(Directory.GetCurrentDirectory(), "CustomDataPointLabels.xlsx");

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CustomDataPointLabelsDemo.Run();
        }
    }
}
