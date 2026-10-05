// Title: Hide Excel chart points whose category label includes “Total” using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook, accesses the first chart, and sets IsVisible = false for any series point whose category cell contains the word "Total" with Aspose.Cells. | Show how to obtain a chart series' CategoryData range and iterate through its cells to filter out categories containing a specific keyword in Aspose.Cells for .NET. | Create a snippet that uses dynamic typing to adjust chart point visibility based on category values when working with Aspose.Cells chart objects.
// Common Searches: asp.net hide chart points with label total using aspose.cells | c# filter chart categories containing the word total in an excel file | how to programmatically exclude total rows from a chart series in aspose.cells | apply category filter to excel chart series with aspose.cells .net | remove total category from chart series using aspose.cells api
// Tags: category data range filtering Aspose.Cells | chart point visibility control C# | dynamic series manipulation Aspose.Cells | total label filtering Aspose.Cells | excel chart series programmatic filtering

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeRange = Aspose.Cells.Range;

// The example loads a workbook, accesses the first chart and its first series, retrieves the series' CategoryData range, iterates through each category cell, and sets the corresponding point's IsVisible property to false when the cell text contains "Total", then saves the updated workbook.
class ChartCategoryFilterExample
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet and its first chart
            Worksheet sheet = workbook.Worksheets[0];
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            Chart chart = sheet.Charts[0];

            // Get the first series of the chart
            Series series = chart.NSeries[0];

            // Use dynamic to access members that may vary between Aspose.Cells versions
            dynamic dynSeries = series;

            // Retrieve the category data range (e.g., "Sheet1!A2:A10")
            string categoryRange = dynSeries.CategoryData;
            if (string.IsNullOrEmpty(categoryRange))
                throw new InvalidOperationException("Series does not have category data defined.");

            // Create a range object for the category cells
            AsposeRange catRange = sheet.Cells.CreateRange(categoryRange);
            if (catRange == null)
                throw new InvalidOperationException($"Unable to create range from '{categoryRange}'.");

            // Iterate through each category cell and set point visibility based on its text
            int pointIndex = 0;
            foreach (Cell cell in catRange)
            {
                string category = cell.StringValue;
                bool isVisible = !category.Contains("Total", StringComparison.OrdinalIgnoreCase);

                // Ensure the point exists before setting visibility
                if (pointIndex < series.Points.Count)
                {
                    dynamic point = series.Points[pointIndex];
                    point.IsVisible = isVisible;
                }

                pointIndex++;
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
