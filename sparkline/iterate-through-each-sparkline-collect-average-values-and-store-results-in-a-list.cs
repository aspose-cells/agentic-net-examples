// Title: Calculate and collect average values of each sparkline in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: how to compute average of sparkline data range with Aspose.Cells C# | iterate through sparkline groups and retrieve numeric cells using Aspose.Cells .NET | extract sparkline data range values and calculate average in C# with Aspose.Cells | store sparkline averages in a List<double> when processing an Excel file | process all worksheets to calculate sparkline metrics using Aspose.Cells
// Tags: calculate sparkline average Aspose.Cells | iterate sparkline groups .NET | read sparkline data range cells | collect sparkline metrics List<double> | process workbook sparklines C#

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, walks through every worksheet, sparkline group, and individual sparkline, reads the numeric cells in each sparkline's DataRange, computes the average (or zero if no numeric cells), adds each average to a List<double>, and saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // List to hold the average value of each sparkline
            List<double> sparklineAverages = new List<double>();

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Use dynamic to avoid compile‑time dependency on Sparkline types
                dynamic sparklineGroups = sheet.SparklineGroups;
                int groupCount = sparklineGroups.Count;

                for (int g = 0; g < groupCount; g++)
                {
                    try
                    {
                        dynamic group = sparklineGroups[g];
                        dynamic sparklineCollection = group.SparklineCollection;
                        int sparkCount = sparklineCollection.Count;

                        for (int s = 0; s < sparkCount; s++)
                        {
                            try
                            {
                                dynamic spark = sparklineCollection[s];
                                // DataRange is of type CellArea
                                CellArea dataRange = spark.DataRange;

                                double sum = 0.0;
                                int count = 0;

                                // Loop through all cells in the data range
                                for (int row = dataRange.StartRow; row <= dataRange.EndRow; row++)
                                {
                                    for (int col = dataRange.StartColumn; col <= dataRange.EndColumn; col++)
                                    {
                                        Cell cell = sheet.Cells[row, col];

                                        // Consider only numeric cells for the average calculation
                                        if (cell.Type == CellValueType.IsNumeric)
                                        {
                                            sum += cell.DoubleValue;
                                            count++;
                                        }
                                    }
                                }

                                // Compute average; if no numeric cells, average is set to 0
                                double average = count > 0 ? sum / count : 0.0;
                                sparklineAverages.Add(average);
                            }
                            catch (Exception exInner)
                            {
                                Console.WriteLine($"Error processing sparkline #{s + 1} in group #{g + 1}: {exInner.Message}");
                            }
                        }
                    }
                    catch (Exception exGroup)
                    {
                        Console.WriteLine($"Error processing sparkline group #{g + 1}: {exGroup.Message}");
                    }
                }
            }

            // Save the workbook (if any modifications were made)
            workbook.Save(outputPath);
            Console.WriteLine($"Processed {sparklineAverages.Count} sparklines. Output saved to {outputPath}.");
        }
        catch (Exception ex)
        {
            // Runtime safety: capture any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
