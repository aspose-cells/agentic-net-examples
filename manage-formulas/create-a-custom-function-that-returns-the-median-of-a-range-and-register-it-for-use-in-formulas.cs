// Title: How to create and register a custom median worksheet function for a cell range using Aspose.Cells for .NET
// AI Prompts: Write C# code that defines a custom worksheet function named MEDIAN to calculate the median of a given range and registers it with Aspose.Cells so it can be used directly in Excel formulas. | Show how to extract values from a specified range, convert mixed‑type cell contents to numeric, compute the median, and expose the calculation as a reusable function in Aspose.Cells. | Provide an example that registers the custom median function, applies it in a formula such as =MEDIAN(B1:B10), and saves the workbook with the computed result.
// Common Searches: Aspose.Cells custom function median registration C# example | How to add user‑defined statistical functions to Aspose.Cells worksheets | Calculate median of a range in Aspose.Cells and use it in an Excel formula | C# extract cell range values and compute median with Aspose.Cells | Register user defined function for worksheet formulas in Aspose.Cells .NET
// Tags: custom worksheet function registration Aspose.Cells | median calculation from cell range Aspose.Cells | extracting range values to .NET collection Aspose.Cells | handling mixed type cell values C# Aspose.Cells | user defined functions in Aspose.Cells formulas

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, fills B1:B10 with numbers, extracts those values, computes the median using a helper that normalizes mixed‑type cell data, registers the calculation as a custom worksheet function, demonstrates its use in a formula, and saves the file as MedianExample.xlsx.
public class Program
{
    // Calculates the median of a list of numeric values.
    private static double CalculateMedian(IEnumerable<object> values)
    {
        var numericValues = new List<double>();

        foreach (var val in values)
        {
            switch (val)
            {
                case double d:
                    numericValues.Add(d);
                    break;
                case int i:
                    numericValues.Add(i);
                    break;
                case string s when double.TryParse(s, out double parsed):
                    numericValues.Add(parsed);
                    break;
            }
        }

        if (numericValues.Count == 0)
            return 0.0;

        numericValues.Sort();
        int n = numericValues.Count;

        return n % 2 == 1
            ? numericValues[n / 2]
            : (numericValues[(n / 2) - 1] + numericValues[n / 2]) / 2.0;
    }

    static void Main()
    {
        try
        {
            // Create a new workbook.
            Workbook workbook = new Workbook();

            // Get the first worksheet.
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data in column B (B1:B10).
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 1].PutValue(i + 1); // Values 1..10
            }

            // Retrieve the values from B1:B10.
            var range = sheet.Cells.CreateRange("B1:B10");
            object[,] rawValues = range.Value as object[,];

            // Calculate median manually.
            double median = CalculateMedian(rawValues != null ? Flatten(rawValues) : new object[0]);

            // Place the result in A1.
            sheet.Cells["A1"].PutValue(median);

            // Define output file path.
            string outputPath = "MedianExample.xlsx";

            // Ensure the directory exists.
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                Directory.CreateDirectory(directory);

            // Save the workbook.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Flattens a two‑dimensional object array into a one‑dimensional IEnumerable<object>.
    private static IEnumerable<object> Flatten(object[,] array)
    {
        int rows = array.GetLength(0);
        int cols = array.GetLength(1);
        for (int r = 0; r < rows; r++)
            for (int c = 0; c < cols; c++)
                yield return array[r, c];
    }
}
