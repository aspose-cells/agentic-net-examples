using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class HistogramBinExample
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample numeric data in column A (A1:A10)
        double[] data = { 5, 12, 18, 22, 27, 33, 41, 49, 55, 63 };
        for (int i = 0; i < data.Length; i++)
        {
            sheet.Cells[i, 0].PutValue(data[i]); // column index 0 = column A
        }

        // Add a histogram chart to the worksheet
        int chartIndex = sheet.Charts.Add(ChartType.Histogram, 5, 2, 25, 15);
        Chart chart = sheet.Charts[chartIndex];

        // Use the data range as the series source
        chart.NSeries.Add("A1:A10", true);

        // Configure manual bins on the value axis
        AxisBins bins = chart.ValueAxis.Bins;
        bins.IsAutomatic = false;   // turn off automatic bin calculation
        bins.Width = 10;            // each bin spans 10 units
        bins.Count = 7;             // total number of bins
        bins.Underflow = 0;         // values below this go to underflow bin
        bins.Overflow = 70;         // values above this go to overflow bin

        // Enumerate the source column and map each value to its bin index
        Console.WriteLine("Value -> Bin Index");
        for (int row = 0; row < data.Length; row++)
        {
            double val = sheet.Cells[row, 0].DoubleValue;
            int binIndex;

            if (val < bins.Underflow)
                binIndex = -1;                     // underflow bin
            else if (val > bins.Overflow)
                binIndex = bins.Count;              // overflow bin
            else
                binIndex = (int)Math.Floor((val - bins.Underflow) / bins.Width);

            Console.WriteLine($"{val} -> {binIndex}");
        }

        // Save the workbook with the configured histogram
        workbook.Save("HistogramBins.xlsx");
    }
}