using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // ------------------------------------------------------------
            // Populate sample numeric data in column B (B2:B11)
            // ------------------------------------------------------------
            double[] sampleValues = { 5, 12, 18, 22, 27, 33, 38, 44, 51, 59 };
            for (int i = 0; i < sampleValues.Length; i++)
            {
                // Row index is i+1 because row 0 is header (optional)
                worksheet.Cells[i + 1, 1].PutValue(sampleValues[i]); // Column index 1 => column B
            }

            // Optional: add a header for clarity
            worksheet.Cells[0, 1].PutValue("Values");

            // ------------------------------------------------------------
            // Enumerate the values in column B and print them (demonstrates enumeration)
            // ------------------------------------------------------------
            Console.WriteLine("Enumerating column B values:");
            for (int row = 1; row <= sampleValues.Length; row++)
            {
                Cell cell = worksheet.Cells[row, 1];
                Console.WriteLine($"Row {row + 1}: {cell.DoubleValue}");
            }

            // ------------------------------------------------------------
            // Add a histogram chart that uses the data in column B
            // ------------------------------------------------------------
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Histogram, 5, 3, 25, 13);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the data series – the range is B2:B11, and we set isVertical = true
            chart.NSeries.Add("B2:B11", true);

            // ------------------------------------------------------------
            // Configure axis bins for the value axis (the X‑axis of a histogram)
            // ------------------------------------------------------------
            Axis valueAxis = chart.ValueAxis;          // X‑axis for histogram values
            AxisBins bins = valueAxis.Bins;

            // Disable automatic bin calculation to define custom bins
            bins.IsAutomatic = false;

            // Define bin properties
            bins.Width = 10;       // Width of each bin
            bins.Count = 6;        // Number of bins (optional, overrides width if needed)
            bins.Overflow = 60;    // Values greater than this go to overflow bin
            bins.Underflow = 0;    // Values less than this go to underflow bin

            // ------------------------------------------------------------
            // Save the workbook to a file
            // ------------------------------------------------------------
            workbook.Save("HistogramWithCustomBins.xlsx");

            Console.WriteLine("Histogram chart created and saved as 'HistogramWithCustomBins.xlsx'.");
        }
    }
}