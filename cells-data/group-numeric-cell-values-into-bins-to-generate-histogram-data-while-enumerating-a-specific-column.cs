// Title: Generate a histogram chart with custom bin width, overflow/underflow settings and enumerate numeric column values using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that fills a worksheet column with numeric values, prints each cell’s value to the console, and adds a histogram chart with custom bin width, count, overflow, and underflow using Aspose.Cells. | Modify the example to read data from a different column range and configure the histogram bins by specifying the number of bins instead of the bin width. | Implement a reusable method that accepts a worksheet, a data range address, and bin parameters (width, count, overflow, underflow) and inserts a histogram chart with those settings.
// Common Searches: aspacells c# create histogram chart with custom bin intervals and overflow | how to enumerate cells in a specific column using Aspose.Cells for .NET | configure overflow and underflow bins for a histogram in Aspose.Cells | custom bin count vs width Aspose.Cells histogram example | generate histogram from column B values Aspose.Cells C#
// Tags: Aspose.Cells histogram custom bins | C# enumerate worksheet column values | Aspose.Cells set bin width | Aspose.Cells overflow underflow bins | Aspose.Cells chart from data range | Aspose.Cells save workbook as xlsx

using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsHistogramExample
{
    // // Creates a new workbook, populates column B (B2:B11) with numeric data, prints each value to the console, adds a histogram chart with custom bin width, count, overflow, and underflow settings, and saves the workbook as 'HistogramWithCustomBins.xlsx'.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate numeric data in column B (B2:B11)
            double[] values = { 5, 12, 18, 22, 27, 33, 38, 45, 51, 60 };
            for (int i = 0; i < values.Length; i++)
            {
                // Row index is i+1 because row 0 is header
                worksheet.Cells[i + 1, 1].PutValue(values[i]);
            }

            // Optional: add a header for the column
            worksheet.Cells[0, 1].PutValue("Values");

            // Enumerate the specific column and output values to console
            Console.WriteLine("Enumerating column B values:");
            for (int row = 1; row <= values.Length; row++)
            {
                Console.WriteLine($"Row {row + 1}: {worksheet.Cells[row, 1].DoubleValue}");
            }

            // Add a histogram chart to the worksheet
            // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
            int chartIndex = worksheet.Charts.Add(ChartType.Histogram, 5, 3, 25, 13);
            Chart chart = worksheet.Charts[chartIndex];

            // Add the data series (numeric values) to the chart
            // The second argument 'true' indicates that the range is a category (X) axis data
            chart.NSeries.Add("B2:B11", true);

            // Configure the bins on the value axis (the axis that holds the numeric data)
            Axis valueAxis = chart.ValueAxis;
            AxisBins bins = valueAxis.Bins;

            // Disable automatic bin calculation to define custom bins
            bins.IsAutomatic = false;

            // Set custom bin properties
            bins.Width = 10;      // Width of each bin
            bins.Count = 6;       // Number of bins (optional, can be omitted if Width is set)
            bins.Overflow = 70;   // Values greater than the last bin go to overflow bin
            bins.Underflow = 0;   // Values less than the first bin go to underflow bin

            // Save the workbook to a file
            workbook.Save("HistogramWithCustomBins.xlsx");

            Console.WriteLine("Histogram chart created and workbook saved as 'HistogramWithCustomBins.xlsx'.");
        }
    }
}
