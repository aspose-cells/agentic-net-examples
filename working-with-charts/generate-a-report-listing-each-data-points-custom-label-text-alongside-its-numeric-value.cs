// Title: Aspose.Cells for .NET – C# Console Report of Custom Chart Data Labels and Their Values
// Description: This C# sample creates a workbook, adds a column chart with three points, assigns custom label texts (Alpha, Beta, Gamma), disables automatic labeling, and prints each custom label together with its numeric Y‑value to the console before saving the file as an Excel workbook.
// Keywords: Aspose.Cells C# chart custom labels | print chart data label values | Aspose.Cells console report | retrieve chart point value .NET | custom data labels Excel chart | Aspose.Cells chart point iteration
// Common Searches: how to display custom data label text with values using Aspose.Cells | Aspose.Cells C# chart point value extraction | print custom chart labels to console Aspose.Cells | C# Aspose.Cells generate data label report
// Developer Intent: List each chart point’s custom label alongside its numeric value in a console output.
// Use Cases: Create an audit log of chart data where custom labels map to business codes. | Export label‑value pairs to a text or CSV file for downstream processing. | Validate that custom labels applied to chart points match expected categories in automated tests.
// AI Prompts: Show how to include the category name in the console output for each data point. | Provide code that writes the label‑value pairs to a CSV file instead of the console. | Explain how to read and display the font style of each data label in the chart.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabelReport
{
    // This C# sample creates a workbook, adds a column chart with three points, assigns custom label texts (Alpha, Beta, Gamma), disables automatic labeling, and prints each custom label together with its numeric Y‑value to the console before saving the file as an Excel workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B2"].PutValue(100);
                sheet.Cells["B3"].PutValue(200);
                sheet.Cells["B4"].PutValue(300);

                // Add a column chart
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set data range for the series
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true; // show numeric value
                series.DataLabels.ShowCategoryName = false;

                // Define custom label texts
                string[] customLabels = { "Alpha", "Beta", "Gamma" };

                // Assign custom text to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    // Disable auto-generated text so we can set our own
                    point.DataLabels.IsAutoText = false;
                    // Set the custom label text
                    point.DataLabels.Text = customLabels[i];
                }

                // Generate the report: list custom label and its numeric value
                Console.WriteLine("Data Point Report:");
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];
                    string label = point.DataLabels.Text;               // custom label
                    double value = Convert.ToDouble(point.YValue);      // numeric value
                    Console.WriteLine($"Label: {label}, Value: {value}");
                }

                // Save the workbook (ensure the directory exists)
                string outputPath = "CustomDataLabelReport.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
