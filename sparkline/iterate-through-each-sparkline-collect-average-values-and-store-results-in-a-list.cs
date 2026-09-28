// Title: Calculate and store the average of each sparkline’s data range in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that uses Aspose.Cells to loop through every SparklineGroup in a worksheet, retrieve each sparkline’s DataRange, compute the numeric average of the cells, and add the result to a List<double>. | Show how to create an Aspose.Range from a sparkline’s DataRange string, iterate over its values, and calculate the mean while handling non‑numeric cells. | Demonstrate printing each sparkline’s average to the console and saving the workbook after processing all sparklines.
// Common Searches: aspnet cells c# iterate sparkline groups get source cells | how to compute mean of sparkline values with Aspose.Cells | extract numeric cells from sparkline data range in .NET | c# example for listing sparkline averages using Aspose.Cells | saving workbook after processing sparkline averages Aspose
// Tags: aspose.cells sparkline average calculation | c# sparkline group iteration | aspose.cells create range from datarange | compute numeric mean of excel sparkline data | store sparkline averages list<double>

using System;
using System.Collections.Generic;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Alias to avoid conflict with System.Range (C# 8.0)
using AsposeRange = Aspose.Cells.Range;

// The example creates a workbook, adds a line sparkline for cells A1:D1, iterates through all sparkline groups, extracts each sparkline's DataRange, computes the numeric average of its values, stores the averages in a List<double>, prints them to the console, and saves the workbook.
class SparklineAverageExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline
            worksheet.Cells["A1"].PutValue(5);
            worksheet.Cells["B1"].PutValue(2);
            worksheet.Cells["C1"].PutValue(1);
            worksheet.Cells["D1"].PutValue(3);

            // Define the location where the sparkline will be placed (E1)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 0,
                StartColumn = 4,
                EndColumn = 4
            };

            // Add a sparkline group (Line type) and a sparkline inside it
            int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line, "A1:D1", false, location);
            SparklineGroup group = worksheet.SparklineGroups[groupIndex];
            group.Sparklines.Add($"{worksheet.Name}!A1:D1", 0, 4);

            // List to hold average values of each sparkline
            List<double> sparklineAverages = new List<double>();

            // Iterate through each sparkline in all groups
            foreach (SparklineGroup sg in worksheet.SparklineGroups)
            {
                foreach (Sparkline sp in sg.Sparklines)
                {
                    // Get the data range string (e.g., "A1:D1")
                    string dataRange = sp.DataRange;

                    // Create a Range object from the data range string
                    AsposeRange range = worksheet.Cells.CreateRange(dataRange);
                    object[,] values = range.Value as object[,];

                    double sum = 0;
                    int count = 0;

                    // Sum numeric values and count them
                    if (values != null)
                    {
                        for (int i = 0; i < values.GetLength(0); i++)
                        {
                            for (int j = 0; j < values.GetLength(1); j++)
                            {
                                if (values[i, j] != null && double.TryParse(values[i, j].ToString(), out double d))
                                {
                                    sum += d;
                                    count++;
                                }
                            }
                        }
                    }

                    // Compute average (0 if no numeric cells)
                    double average = count > 0 ? sum / count : 0;
                    sparklineAverages.Add(average);
                }
            }

            // Output the averages to the console
            for (int i = 0; i < sparklineAverages.Count; i++)
            {
                Console.WriteLine($"Sparkline {i} average: {sparklineAverages[i]}");
            }

            // Save the workbook
            string outputPath = "SparklineAverages.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
