// Title: C# script to extract the "Other" label from every pie or 3‑D pie chart in an Excel workbook using Aspose.Cells
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, loops through all worksheets, identifies pie and 3‑D pie charts, reads the XValues range of the first series, and prints the category text where the label equals "Other". | Extend the script to also retrieve the numeric value associated with the "Other" slice and display it alongside the label for each matching chart. | Add robust error handling that logs warnings for charts without XValues, missing "Other" categories, or empty ranges, and ensures processing continues for remaining charts.
// Common Searches: how to read pie chart category names from Excel using Aspose.Cells in C# | C# extract specific slice label from Excel pie chart with Aspose.Cells | Aspose.Cells get XValues range of chart series for localization checks | iterate through all charts in a workbook and find "Other" label using Aspose.Cells | read category values of pie chart series in .NET Aspose.Cells
// Tags: Aspose.Cells read pie chart XValues | C# extract chart category label | Excel pie chart "Other" slice detection | iterate worksheets charts Aspose.Cells | localization quality check Excel chart labels

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program loads an Excel workbook, iterates each worksheet and its charts, processes only pie and 3‑D pie charts, reads the XValues range of the first series, flattens the category array, searches for the "Other" label, and outputs the worksheet name, chart name, and the found label text.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook containing the pie charts
            var workbook = new Workbook(inputPath);

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all charts on the worksheet
                foreach (Chart chart in sheet.Charts)
                {
                    // Process only pie charts (2D or 3D)
                    if (chart.Type == ChartType.Pie || chart.Type == ChartType.Pie3D)
                    {
                        // Ensure the chart has at least one series
                        if (chart.NSeries.Count > 0)
                        {
                            // Use the first series (commonly the data source for pie charts)
                            Series series = chart.NSeries[0];

                            // XValues holds the address of the category range (e.g., "A1:A5")
                            string xRange = series.XValues;
                            if (string.IsNullOrEmpty(xRange))
                                continue;

                            // Retrieve the actual category values from the worksheet
                            object[,] categoryArray = sheet.Cells.CreateRange(xRange).Value as object[,];
                            if (categoryArray == null)
                                continue;

                            // Flatten the 2‑D array to a 1‑D array for easier processing
                            int rows = categoryArray.GetLength(0);
                            int cols = categoryArray.GetLength(1);
                            object[] categories = new object[rows * cols];
                            int idx = 0;
                            for (int r = 0; r < rows; r++)
                                for (int c = 0; c < cols; c++)
                                    categories[idx++] = categoryArray[r, c];

                            // Locate the index of the category named "Other"
                            int otherIndex = -1;
                            for (int i = 0; i < categories.Length; i++)
                            {
                                if (categories[i] != null && categories[i].ToString() == "Other")
                                {
                                    otherIndex = i;
                                    break;
                                }
                            }

                            // If "Other" was found, output its label (category name)
                            if (otherIndex >= 0 && otherIndex < categories.Length)
                            {
                                string labelText = categories[otherIndex]?.ToString() ?? string.Empty;
                                Console.WriteLine($"Worksheet: {sheet.Name}, Chart: {chart.Name}, \"Other\" label: {labelText}");
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the program from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
