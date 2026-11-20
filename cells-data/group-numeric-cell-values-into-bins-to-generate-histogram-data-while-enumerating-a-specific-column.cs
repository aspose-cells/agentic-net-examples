// Title: Create a histogram of numeric values from a specific Excel column using Aspose.Cells for .NET (C#)
// Common Searches: c# Aspose.Cells read numeric values from a column and build histogram | how to calculate bin counts for Excel data with Aspose.Cells in .NET | generate histogram worksheet from column B using Aspose.Cells C# example | Aspose.Cells bin width calculation and histogram output to new sheet | group Excel column numbers into bins with Aspose.Cells library
// Tags: Aspose.Cells read numeric column values | Aspose.Cells calculate histogram bins | Aspose.Cells write histogram worksheet | C# binning numeric data in Excel with Aspose | Excel histogram generation using Aspose.Cells .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace HistogramApp
{
    // The example loads an Excel workbook, extracts numeric values from a chosen column, computes min/max and bin width for a fixed number of bins, counts values per bin, creates a new worksheet named "Histogram" with range and count columns, and saves the updated file.
    class HistogramGenerator
    {
        static void Main()
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load workbook (empty password works for both protected and unprotected files)
                LoadOptions loadOptions = new LoadOptions { Password = "" };
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Get the first worksheet
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Column to analyze (0‑based). Example: column B => index 1
                int targetColumnIndex = 1;

                // Collect numeric values from the column
                List<double> values = new List<double>();
                int startRow = sourceSheet.Cells.MinRow;
                int endRow = sourceSheet.Cells.MaxRow;

                for (int row = startRow; row <= endRow; row++)
                {
                    Cell cell = sourceSheet.Cells[row, targetColumnIndex];
                    if (cell.Type == CellValueType.IsNumeric)
                    {
                        values.Add(cell.DoubleValue);
                    }
                }

                if (values.Count == 0)
                {
                    Console.WriteLine("No numeric data found in the specified column.");
                    return;
                }

                // Histogram parameters
                int binCount = 5;

                // Determine min and max values
                double min = double.MaxValue;
                double max = double.MinValue;
                foreach (double v in values)
                {
                    if (v < min) min = v;
                    if (v > max) max = v;
                }

                // Compute bin width
                double range = max - min;
                double binWidth = range / binCount;
                if (binWidth == 0) binWidth = 1; // avoid division by zero for uniform data

                // Count values per bin
                int[] binCounts = new int[binCount];
                foreach (double v in values)
                {
                    int binIndex = (int)Math.Floor((v - min) / binWidth);
                    if (binIndex >= binCount) binIndex = binCount - 1; // include max value
                    binCounts[binIndex]++;
                }

                // Add a new worksheet for the histogram
                Worksheet histogramSheet = workbook.Worksheets[workbook.Worksheets.Add()];
                histogramSheet.Name = "Histogram";

                // Write headers
                histogramSheet.Cells[0, 0].PutValue("Bin Range");
                histogramSheet.Cells[0, 1].PutValue("Count");

                // Populate histogram data
                for (int i = 0; i < binCount; i++)
                {
                    double binStart = min + i * binWidth;
                    double binEnd = binStart + binWidth;
                    string rangeLabel = $"{binStart:F2} - {binEnd:F2}";
                    histogramSheet.Cells[i + 1, 0].PutValue(rangeLabel);
                    histogramSheet.Cells[i + 1, 1].PutValue(binCounts[i]);
                }

                // Save the workbook with the new sheet
                workbook.Save(outputPath);
                Console.WriteLine($"Histogram saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
