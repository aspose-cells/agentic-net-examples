// Title: C# Aspose.Cells: Create a Histogram with Manual Bin Width, Count, Overflow & Underflow
// Description: This example builds a new workbook, writes numeric values to column B, adds a Histogram chart, and configures the ValueAxis bins manually (fixed width, specific count, overflow at 70, underflow at 0) before saving as HistogramWithBins.xlsx.
// Keywords: Aspose.Cells histogram C# | manual histogram bins | custom bin width Aspose.Cells | overflow underflow bins Excel | C# generate histogram chart | Aspose.Cells chart axis bins | Excel data binning .NET
// Common Searches: Aspose.Cells set custom bins for histogram | C# histogram chart overflow underflow | manual bin width Aspose.Cells chart | how to define bin count in Aspose.Cells histogram | create histogram from column data using Aspose.Cells
// Developer Intent: Programmatically produce a histogram chart with user‑defined bin intervals and outlier handling for numeric data in an Excel worksheet.
// Use Cases: Generate sales distribution reports with 10‑unit bins and separate outliers beyond the 0‑70 range. | Visualize test‑score frequencies, applying a fixed bin size while capturing extreme scores in dedicated overflow/underflow bins. | Automate monthly analytics by inserting a pre‑configured histogram into a workbook and exporting it as part of a larger reporting pipeline.
// AI Prompts: Modify the sample to read numbers from an existing Excel file and apply the same manual bin settings. | Add data labels showing the count of items in each histogram bin using Aspose.Cells C#. | Export the configured histogram chart to PNG or JPEG format after setting custom bins.

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;

// This example builds a new workbook, writes numeric values to column B, adds a Histogram chart, and configures the ValueAxis bins manually (fixed width, specific count, overflow at 70, underflow at 0) before saving as HistogramWithBins.xlsx.
class HistogramBinExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Sample numeric data placed in column B (index 1)
        double[] data = { 5, 12, 18, 22, 27, 33, 40, 45, 52, 60 };
        for (int i = 0; i < data.Length; i++)
        {
            // Row index is i+1 because row 0 is header (optional)
            sheet.Cells[i + 1, 1].PutValue(data[i]);
        }

        // Add a histogram chart to the worksheet
        // Parameters: chart type, upper-left row, upper-left column, lower-right row, lower-right column
        int chartIdx = sheet.Charts.Add(ChartType.Histogram, 5, 3, 25, 15);
        Chart chart = sheet.Charts[chartIdx];

        // Define the data series for the histogram (B2:B11)
        chart.NSeries.Add("B2:B11", true);

        // Configure manual bins on the value axis
        AxisBins bins = chart.ValueAxis.Bins;
        bins.IsAutomatic = false;   // Disable automatic bin calculation
        bins.Width = 10;             // Width of each bin
        bins.Count = 6;              // Number of bins
        bins.Overflow = 70;          // Values greater than 70 go to overflow bin
        bins.Underflow = 0;          // Values less than 0 go to underflow bin
        bins.IsByCategory = false;  // Group by numeric range, not by category

        // Save the workbook with the histogram chart
        workbook.Save("HistogramWithBins.xlsx");
    }
}
