// Title: Create histogram bins from numeric values in a specific Excel column using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells read numeric values from a column and build histogram bins | How to calculate bin counts for a column of numbers using Aspose.Cells .NET | Aspose.Cells generate histogram worksheet from source data in C# | Enumerate Excel column values and group into equal‑width bins with Aspose.Cells | Create histogram data table in a new sheet using Aspose.Cells in C#
// Tags: numeric column extraction Aspose.Cells C# | equal-width bin calculation Aspose.Cells | histogram worksheet creation Aspose.Cells | bin count aggregation Aspose.Cells .NET | Excel data binning Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace HistogramGenerator
{
    // The example loads an existing Excel workbook, extracts all numeric values from a chosen column, determines the data range, divides it into a configurable number of equal‑width bins, counts the values per bin, creates a new worksheet named "Histogram" with columns for bin start, bin end, and count, writes the histogram data, and saves the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Input and output file paths
            string inputPath = @"C:\Data\SourceWorkbook.xlsx";
            string outputPath = @"C:\Data\HistogramWorkbook.xlsx";

            // Load the workbook (lifecycle rule)
            Workbook workbook = new Workbook(inputPath);

            // Choose the worksheet that contains the data
            Worksheet dataSheet = workbook.Worksheets[0]; // first sheet

            // Column index to enumerate (0 = A, 1 = B, etc.)
            int targetColumnIndex = 0;

            // Collect numeric values from the target column
            List<double> values = new List<double>();
            // Assuming data starts at row 0 and continues until an empty cell is found
            for (int row = 0; row <= dataSheet.Cells.MaxDataRow; row++)
            {
                var cell = dataSheet.Cells[row, targetColumnIndex];
                if (cell.Value == null) continue;

                // Try to parse the cell value as double
                if (double.TryParse(cell.Value.ToString(), out double numericValue))
                {
                    values.Add(numericValue);
                }
            }

            if (values.Count == 0)
            {
                Console.WriteLine("No numeric data found in the specified column.");
                return;
            }

            // Determine histogram parameters
            int numberOfBins = 10; // you can adjust this as needed
            double min = double.MaxValue;
            double max = double.MinValue;

            foreach (double v in values)
            {
                if (v < min) min = v;
                if (v > max) max = v;
            }

            // Edge case: all values are the same
            if (min == max)
            {
                max = min + 1;
            }

            double binSize = (max - min) / numberOfBins;

            // Initialize bin counters
            int[] binCounts = new int[numberOfBins];

            // Assign each value to a bin
            foreach (double v in values)
            {
                int binIndex = (int)((v - min) / binSize);
                // Ensure the max value falls into the last bin
                if (binIndex == numberOfBins) binIndex = numberOfBins - 1;
                binCounts[binIndex]++;
            }

            // Create a new worksheet to store histogram data
            Worksheet histogramSheet = workbook.Worksheets.Add("Histogram");

            // Write header
            histogramSheet.Cells[0, 0].PutValue("Bin Start");
            histogramSheet.Cells[0, 1].PutValue("Bin End");
            histogramSheet.Cells[0, 2].PutValue("Count");

            // Populate histogram rows
            for (int i = 0; i < numberOfBins; i++)
            {
                double binStart = min + i * binSize;
                double binEnd = binStart + binSize;

                histogramSheet.Cells[i + 1, 0].PutValue(binStart);
                histogramSheet.Cells[i + 1, 1].PutValue(binEnd);
                histogramSheet.Cells[i + 1, 2].PutValue(binCounts[i]);
            }

            // Save the workbook (lifecycle rule)
            workbook.Save(outputPath);
            Console.WriteLine($"Histogram data saved to {outputPath}");
        }
    }
}
