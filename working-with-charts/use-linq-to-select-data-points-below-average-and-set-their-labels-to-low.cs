// Title: Aspose.Cells C# – Label chart points below the average as “Low” using LINQ
// Description: Creates a workbook with categories and values, adds a column chart, computes the average of the value range, and uses a LINQ query to label every point whose source value is below the average with the text "Low". The workbook is saved as BelowAverageLabelDemo.xlsx.
// Keywords: Aspose.Cells | C# chart labeling | LINQ data points | below average label | column chart Aspose.Cells | custom data labels | average calculation Excel | Aspose.Cells example
// Common Searches: Aspose.Cells label points below average | C# LINQ chart data labels | how to mark low values in Aspose.Cells chart | set custom label for chart points Aspose.Cells | LINQ select chart points below mean
// Developer Intent: Apply a LINQ query to identify chart points whose values are under the calculated average and assign the label "Low" to those points in an Aspose.Cells workbook.
// Use Cases: Highlight under‑performing product categories in a sales chart. | Flag sensor readings that fall below the mean in a monitoring dashboard. | Automatically annotate metrics that are lower than the overall average in a financial report.
// AI Prompts: Rewrite the sample to replace the for‑loop with a LINQ expression that selects points below the average and sets DataLabels.Text to "Low". | Provide a C# method that calculates the average of a range, uses LINQ to find indices of values below that average, and updates the corresponding chart points with a custom label in Aspose.Cells. | Show how to add error handling for missing series while applying LINQ to label low values in a column chart.

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExamples
{
    // Creates a workbook with categories and values, adds a column chart, computes the average of the value range, and uses a LINQ query to label every point whose source value is below the average with the text "Low". The workbook is saved as BelowAverageLabelDemo.xlsx.
    public class BelowAverageLabelDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
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

            // Populate sample data
            // Column A – categories
            // Column B – numeric values
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
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["B6"].PutValue(5);

            // Add a column chart
            int chartIndex = sheet.Charts.Add(ChartType.Column, 7, 0, 25, 8);
            Chart chart = sheet.Charts[chartIndex];

            // Set the data source for the chart
            chart.NSeries.Add("B2:B6", true);
            chart.NSeries.CategoryData = "A2:A6";

            // Calculate the average value from the worksheet data (B2:B6)
            double sum = 0;
            int count = 0;
            for (int row = 1; row <= 5; row++) // rows 2 to 6 (zero‑based index)
            {
                double val = sheet.Cells[row, 1].DoubleValue; // column B (index 1)
                sum += val;
                count++;
            }
            double average = count > 0 ? sum / count : 0;

            // Get all points of the first series
            ChartPointCollection points = chart.NSeries[0].Points;

            // Set label "Low" for points whose source value is below the average
            for (int i = 0; i < points.Count; i++)
            {
                double cellValue = sheet.Cells[i + 1, 1].DoubleValue; // B2, B3, ...
                if (cellValue < average)
                {
                    points[i].DataLabels.ShowValue = true;
                    points[i].DataLabels.Text = "Low";
                }
            }

            // Save the workbook
            try
            {
                workbook.Save("BelowAverageLabelDemo.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save workbook: {ex.Message}");
            }
        }
    }
}
