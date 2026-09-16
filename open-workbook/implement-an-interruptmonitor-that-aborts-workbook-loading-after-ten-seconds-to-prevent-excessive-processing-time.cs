// Title: Abort Aspose.Cells workbook loading after 10 seconds with a custom LoadFilter in C#
// AI Prompts: Write a C# LoadFilter that tracks elapsed time and throws an OperationCanceledException once a specified timeout is exceeded, then apply it via LoadOptions when opening a Workbook. | Show how to catch the timeout exception, log the sheet, row, and column indices at the moment of abort, and optionally save the partially loaded workbook. | Create a reusable method that accepts a timeout value and returns LoadOptions configured with the timeout LoadFilter for any Aspose.Cells workbook load.
// Common Searches: c# set timeout for Aspose.Cells workbook load to prevent long processing | how to cancel loading of a large Excel file with Aspose.Cells after 10 seconds | using LoadFilter to interrupt Aspose.Cells workbook loading based on elapsed time | exception handling for OperationCanceledException during Aspose.Cells load | best practice for limiting Excel load time with Aspose.Cells in .NET
// Tags: loadfilter timeout Aspose.Cells | abort workbook load C# | operationcanceledexception Aspose.Cells | loadoptions custom timeout filter | prevent excessive workbook load time

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// Custom load filter that aborts loading after a specified time interval.
// Demonstrates a custom TimeOutLoadFilter derived from LoadFilter that checks elapsed time on each worksheet, cell, and chart load, throwing an OperationCanceledException after 10 seconds. Shows how to attach the filter to LoadOptions, load a workbook with a timeout, and handle the timeout and other exceptions.
class TimeOutLoadFilter : LoadFilter
{
    private readonly Stopwatch _stopwatch;
    private readonly TimeSpan _maxDuration;

    public TimeOutLoadFilter(TimeSpan maxDuration)
    {
        _maxDuration = maxDuration;
        _stopwatch = Stopwatch.StartNew();
    }

    // Throws an exception if the allowed time has been exceeded.
    private void CheckTimeout()
    {
        if (_stopwatch.Elapsed > _maxDuration)
            throw new OperationCanceledException(
                $"Workbook loading timed out after {_maxDuration.TotalSeconds} seconds.");
    }

    // Called for each worksheet during loading.
    public bool ShouldLoadWorksheet(int sheetIndex)
    {
        CheckTimeout();
        return true; // Load the worksheet.
    }

    // Called for each cell during loading.
    public bool ShouldLoadCell(int sheetIndex, int row, int column)
    {
        CheckTimeout();
        return true; // Load the cell.
    }

    // Called for each chart during loading.
    public bool ShouldLoadChart(int sheetIndex, int chartIndex)
    {
        CheckTimeout();
        return true; // Load the chart.
    }
}

// Example usage.
class Program
{
    static void Main()
    {
        const string inputPath = "LargeWorkbook.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: File \"{inputPath}\" not found.");
            return;
        }

        // Create load options and attach the timeout filter (10 seconds).
        LoadOptions loadOptions = new LoadOptions
        {
            LoadFilter = new TimeOutLoadFilter(TimeSpan.FromSeconds(10))
        };

        Workbook workbook = null;

        try
        {
            // Load the workbook with the timeout monitoring.
            workbook = new Workbook(inputPath, loadOptions);
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (OperationCanceledException ex)
        {
            // Loading was aborted due to timeout.
            Console.WriteLine("Loading aborted: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Handle other possible loading errors.
            Console.WriteLine("Error loading workbook: " + ex.Message);
        }

        // Optionally, save the partially loaded workbook or perform further processing.
        // if (workbook != null)
        //     workbook.Save("PartialWorkbook.xlsx");
    }
}
