// Title: Extract category axis labels from an Aspose.Cells chart and write them to a new worksheet column in C#
// Common Searches: C# Aspose.Cells get chart category axis labels and save to another sheet | how to copy Excel chart axis labels to a separate worksheet using Aspose.Cells | Aspose.Cells extract chart NSeries.CategoryData values to a new worksheet column | store chart category labels in a new sheet with Aspose.Cells .NET | retrieve chart axis label strings and export to worksheet using Aspose.Cells
// Tags: Aspose.Cells chart category axis extraction | write extracted labels to separate worksheet | C# NSeries.CategoryData range handling | Aspose.Cells OoxmlSaveOptions directory creation | populate AxisLabels worksheet with chart data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AxisLabelExtractor
{
    // The example creates a workbook, adds sample data, builds a column chart, reads the chart's CategoryData range to obtain the category axis label strings, creates a new worksheet named "AxisLabels", writes each label into the first column, and saves the workbook using OoxmlSaveOptions with automatic directory creation.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Prepare output directory
                string outputDir = Path.Combine(Directory.GetCurrentDirectory(), "Output");
                Directory.CreateDirectory(outputDir);

                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet dataSheet = workbook.Worksheets[0];
                dataSheet.Name = "Data";

                // Populate sample data: categories in column A, values in column B
                string[] categories = { "Q1", "Q2", "Q3", "Q4" };
                double[] values = { 1200.5, 1500.75, 1100.0, 1800.25 };

                for (int i = 0; i < categories.Length; i++)
                {
                    dataSheet.Cells[i + 1, 0].PutValue(categories[i]); // A2:A5
                    dataSheet.Cells[i + 1, 1].PutValue(values[i]);    // B2:B5
                }

                // Add headers
                dataSheet.Cells[0, 0].PutValue("Quarter");
                dataSheet.Cells[0, 1].PutValue("Revenue");

                // Create a column chart that uses the data range
                int chartIndex = dataSheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
                Chart chart = dataSheet.Charts[chartIndex];
                chart.Title.Text = "Quarterly Revenue";

                // Set the data source for the series (values) and category data (labels)
                chart.NSeries.Add("B2:B5", true);
                chart.NSeries.CategoryData = "A2:A5";

                // -----------------------------------------------------------------
                // Retrieve axis label strings (category axis labels) from the chart
                // -----------------------------------------------------------------
                string categoryRange = chart.NSeries.CategoryData; // e.g., "A2:A5"
                if (string.IsNullOrEmpty(categoryRange))
                {
                    Console.WriteLine("The chart does not have category data defined.");
                    return;
                }

                // Create a range object for the category data
                Aspose.Cells.Range catRange = dataSheet.Cells.CreateRange(categoryRange);
                int labelCount = catRange.RowCount;

                // Prepare a new worksheet to store the extracted labels
                int labelSheetIndex = workbook.Worksheets.Add();
                Worksheet labelSheet = workbook.Worksheets[labelSheetIndex];
                labelSheet.Name = "AxisLabels";

                // Write a header
                labelSheet.Cells[0, 0].PutValue("Axis Label");

                // Copy each label into the new sheet
                for (int i = 0; i < labelCount; i++)
                {
                    Cell srcCell = catRange[i, 0];
                    string labelText = srcCell.StringValue;
                    labelSheet.Cells[i + 1, 0].PutValue(labelText);
                }

                // Save the workbook with both the chart and the extracted labels
                string outputPath = Path.Combine(outputDir, "AxisLabelsDemo.xlsx");

                // Use OoxmlSaveOptions to enable automatic directory creation
                OoxmlSaveOptions saveOptions = new OoxmlSaveOptions
                {
                    CreateDirectory = true
                };
                workbook.Save(outputPath, saveOptions);

                Console.WriteLine($"Workbook saved to: {outputPath}");
                Console.WriteLine("Extracted axis labels have been written to the 'AxisLabels' worksheet.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
