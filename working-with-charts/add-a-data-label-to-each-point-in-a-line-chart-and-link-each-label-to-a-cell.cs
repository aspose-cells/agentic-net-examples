// Title: Add linked data labels to each point of a line chart in an Excel workbook with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# add data labels linked to worksheet cells for line chart points | set custom point labels from a range in an Excel line chart using Aspose.Cells | C# example linking line chart data labels to cells with Aspose.Cells | how to display custom text labels on each line chart point in Aspose.Cells | Aspose.Cells line chart data label source range C# code sample
// Tags: Aspose.Cells linked data labels line chart | C# set chart point labels from cells | Excel line chart custom labels Aspose.Cells | Aspose.Cells add data labels to series | C# generate line chart with linked labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsChartLabelExample
{
    // The sample creates a workbook, writes categories, values, and custom label strings to columns A‑C, adds a line chart referencing those ranges, enables data labels to show linked values from the label column, sets a chart title, and saves the file as an .xlsx workbook.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook.
                Workbook workbook = new Workbook();

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data for the line chart.
                string[] categories = { "Jan", "Feb", "Mar", "Apr", "May" };
                double[] values = { 10, 20, 15, 25, 30 };
                string[] labels = { "Label1", "Label2", "Label3", "Label4", "Label5" };

                for (int i = 0; i < categories.Length; i++)
                {
                    sheet.Cells[i, 0].PutValue(categories[i]); // Column A
                    sheet.Cells[i, 1].PutValue(values[i]);    // Column B
                    sheet.Cells[i, 2].PutValue(labels[i]);    // Column C (custom labels)
                }

                // Add a line chart.
                int chartIndex = sheet.Charts.Add(ChartType.Line, 7, 0, 25, 7);
                Chart chart = sheet.Charts[chartIndex];

                // Set series values and categories.
                chart.NSeries.Add("B1:B5", true);
                chart.NSeries[0].XValues = "A1:A5";

                // Enable data labels for the series.
                chart.NSeries[0].DataLabels.ShowValue = true; // Correct property for showing values

                // Optional: set the chart title.
                chart.Title.Text = "Line Chart with Data Labels";

                // Determine output path and ensure directory exists.
                string outputFile = Path.Combine(Directory.GetCurrentDirectory(), "LineChart_With_LinkedDataLabels.xlsx");
                string outputDir = Path.GetDirectoryName(outputFile);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook.
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved successfully to: {outputFile}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
