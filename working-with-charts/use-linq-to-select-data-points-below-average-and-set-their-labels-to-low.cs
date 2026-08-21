// Title: Aspose.Cells .NET – Label chart points below average as “Low” using LINQ
// Description: Creates a workbook, adds a column chart, computes the series average with LINQ, and sets a custom data label "Low" on every point whose value is below the average before saving the file.
// Keywords: Aspose.Cells | C# chart data labels | LINQ average | label points below average | column chart Aspose | Excel automation .NET | custom data label | highlight low values | chart point labeling | Aspose.Cells example
// Common Searches: Aspose.Cells label low values chart | C# set custom data label for chart points | calculate average of series with LINQ Aspose.Cells | how to mark points below average in Excel chart using Aspose | Aspose.Cells chart point labeling tutorial
// Developer Intent: Add a “Low” label to chart points whose values are below the series average.
// Use Cases: Mark under‑performing sales categories in a column chart. | Automatically flag KPI values that fall under the average in financial reports. | Generate Excel dashboards that highlight data points below a computed threshold. | Create printable reports where low measurements are clearly identified.
// AI Prompts: Generate C# code with Aspose.Cells that computes the average of a column‑chart series using LINQ and sets the DataLabels of points below the average to "Low". | Show how to read chart series values without GetValueArray, calculate the mean, and apply a custom label to each point under the mean in Aspose.Cells. | Extend the example to also label points above the average as "High" while preserving the existing "Low" labels.

using System;
using System.IO;
using System.Linq;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Creates a workbook, adds a column chart, computes the series average with LINQ, and sets a custom data label "Low" on every point whose value is below the average before saving the file.
    public class BelowAverageLabelDemo
    {
        public static void Main()
        {
            try
            {
                Run();
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the chart (values in column B, categories in column A)
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("A");
            sheet.Cells["A3"].PutValue("B");
            sheet.Cells["A4"].PutValue("C");
            sheet.Cells["A5"].PutValue("D");
            sheet.Cells["A6"].PutValue("E");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["B3"].PutValue(25);
            sheet.Cells["B4"].PutValue(15);
            sheet.Cells["B5"].PutValue(30);
            sheet.Cells["B6"].PutValue(5);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Retrieve the numeric values from the source range without using GetValueArray (not available in some versions)
            double[] seriesValues;
            try
            {
                // Rows 2‑6 correspond to indices 1‑5 (zero‑based)
                int startRow = 1;
                int endRow = 5;
                seriesValues = new double[endRow - startRow + 1];
                for (int r = startRow; r <= endRow; r++)
                {
                    object cellValue = sheet.Cells[r, 1].Value; // Column B (index 1)
                    seriesValues[r - startRow] = Convert.ToDouble(cellValue);
                }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException("Failed to read chart data values.", ex);
            }

            // Calculate the average of the series values
            double average = seriesValues.Average();

            // Iterate through points and label those below average
            ChartPointCollection points = chart.NSeries[0].Points;
            for (int i = 0; i < points.Count; i++)
            {
                if (seriesValues[i] < average)
                {
                    ChartPoint point = points[i];
                    point.DataLabels.ShowValue = true; // ensure the label is displayed
                    point.DataLabels.Text = "Low";     // custom label text
                }
            }

            // Save the workbook
            string outputPath = "BelowAverageLabelDemo.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath);
        }
    }
}
