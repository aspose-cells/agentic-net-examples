// Title: C# script to extract the “Other” slice label from pie charts in an Excel workbook with Aspose.Cells
// Common Searches: how to get the 'Other' slice label from an Excel pie chart using Aspose.Cells in C# | C# extract category names from pie chart series with Aspose.Cells | read pie chart data point labels in an .xlsx file using Aspose.Cells | Aspose.Cells locate specific slice label in 2D or 3D pie chart | C# code to iterate Excel chart series and find 'Other' category | extract localization label from pie chart in Aspose.Cells workbook
// Tags: Aspose.Cells extract pie chart slice label | C# iterate Excel chart series values | find specific category in Excel pie chart | read pie chart data points with Aspose.Cells | localization quality check Excel chart labels | process 2D and 3D pie charts using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook with Aspose.Cells, walks through each worksheet and its charts, filters for 2‑D or 3‑D pie charts, and iterates over each series and data point to locate a category named "Other". When the slice is found, the script prints the worksheet name, chart title, and the label text (optionally including the point value). Comprehensive error handling reports missing files or chart‑specific issues without stopping the entire run.
class ExtractOtherLabelFromPieCharts
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts in the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Process only pie charts (2‑D or 3‑D)
                    if (chart.Type == ChartType.Pie || chart.Type == ChartType.Pie3D)
                    {
                        try
                        {
                            // Iterate through each series in the chart using dynamic to avoid compile‑time type dependencies
                            foreach (dynamic series in chart.NSeries)
                            {
                                // Get the category (X) values source
                                dynamic xValues = series.XValues;

                                // Get the number of data points in the series
                                int pointCount = series.Points.Count;

                                // Loop through each data point to find the "Other" label
                                for (int i = 0; i < pointCount; i++)
                                {
                                    // Retrieve the category name for the current point
                                    object categoryObj = xValues.GetValue(i);
                                    string category = categoryObj?.ToString() ?? string.Empty;

                                    // Check if this point is the "Other" slice
                                    if (category.Equals("Other", StringComparison.OrdinalIgnoreCase))
                                    {
                                        // Access the data point (optional, e.g., to get its value)
                                        dynamic point = series.Points[i];

                                        // Construct the label text (here we use the category name)
                                        string labelText = category;

                                        // Example: include the point value if needed
                                        // labelText = $"{category} ({point.Value})";

                                        Console.WriteLine($"Worksheet: {sheet.Name}, Chart: {chart.Title?.Text ?? "Untitled"}");
                                        Console.WriteLine($"Found \"Other\" label: {labelText}");
                                    }
                                }
                            }
                        }
                        catch (Exception exChart)
                        {
                            // Handle errors specific to chart processing without stopping the whole run
                            Console.WriteLine($"Error processing chart \"{chart.Title?.Text ?? "Untitled"}\" on sheet \"{sheet.Name}\": {exChart.Message}");
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors during processing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
