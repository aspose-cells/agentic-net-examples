// Title: C# – Assign custom text to each chart data point with DataPoint.DataLabels.Text in Aspose.Cells for .NET
// Description: Creates a workbook, adds a column chart, enables data labels, turns off automatic label generation, reads category names from column A, converts each point's Y‑value, and sets a custom label in the form "Category: Value units" using the DataLabels.Text property before saving the file.
// Keywords: Aspose.Cells | C# | .NET | chart data labels | custom data point label | DataPoint.DataLabels.Text | disable auto text | Excel chart label programmatically | column chart Aspose.Cells | Aspose.Cells API example
// Common Searches: Aspose.Cells set custom data label text | DataPoint.DataLabels.Text C# example | disable automatic chart label Aspose.Cells | how to format chart point labels in .NET | custom label for each column in Excel chart
// Developer Intent: Programmatically assign a unique label to every data point in an Excel chart using Aspose.Cells for .NET.
// Use Cases: Generate financial or scientific reports where each column shows a combined category and value label. | Create dashboards that require non‑auto‑generated, unit‑specific labels for clearer data interpretation. | Build dynamic Excel charts that automatically reflect worksheet changes in their point labels.
// AI Prompts: Show C# code that sets a custom text for each chart point in Aspose.Cells using DataLabels.Text and disables auto text. | Provide an Aspose.Cells example that reads category names from a worksheet and formats point labels as "Category: Value units". | Explain how to programmatically customize chart data labels in .NET so they display both the category and its numeric value.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsCustomDataLabels
{
    // Creates a workbook, adds a column chart, enables data labels, turns off automatic label generation, reads category names from column A, converts each point's Y‑value, and sets a custom label in the form "Category: Value units" using the DataLabels.Text property before saving the file.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the chart
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Alpha");
                sheet.Cells["A3"].PutValue("Beta");
                sheet.Cells["A4"].PutValue("Gamma");

                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["B2"].PutValue(150);
                sheet.Cells["B3"].PutValue(250);
                sheet.Cells["B4"].PutValue(350);

                // Add a column chart to the worksheet
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 20, 12);
                Chart chart = sheet.Charts[chartIndex];

                // Set the data range for the series and categories
                chart.NSeries.Add("B2:B4", true);
                chart.NSeries.CategoryData = "A2:A4";

                // Enable data labels for the first series
                Series series = chart.NSeries[0];
                series.DataLabels.ShowValue = true; // Show the numeric value as part of the label

                // Assign custom text to each data point
                for (int i = 0; i < series.Points.Count; i++)
                {
                    ChartPoint point = series.Points[i];

                    // Disable automatic text generation so we can set custom text
                    point.DataLabels.IsAutoText = false;

                    // Retrieve category name from column A
                    string category = sheet.Cells[i + 2, 0].StringValue; // A2, A3, A4

                    // YValue is returned as object; convert to double safely
                    double value = Convert.ToDouble(point.YValue);

                    // Set custom label text
                    point.DataLabels.Text = $"{category}: {value} units";
                }

                // Determine output file path and ensure directory exists
                string outputPath = "CustomDataLabels.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
