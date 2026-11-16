// Title: Implement a rolling line chart that automatically shifts its data window each minute using Aspose.Cells for .NET
// Common Searches: how to use System.Timers.Timer to update an Aspose.Cells chart every minute in C# | Aspose.Cells rolling chart example with sliding data window | C# code to shift chart series range dynamically in Aspose.Cells | update line chart data range periodically using Aspose.Cells .NET | create a moving window chart in Excel with Aspose.Cells and timer
// Tags: timer-driven chart series update Aspose.Cells | sliding window line chart .NET | dynamic Excel chart range Aspose.Cells | periodic workbook save for chart refresh | C# Aspose.Cells rolling data window

using System;
using System.IO;
using System.Timers;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a Workbook, fills 30 columns with random numbers, adds a line chart, and defines a 10‑column rolling window. A System.Timers.Timer set to 60 000 ms increments the start column, rebuilds the chart series with the new range, updates the chart title, and saves the workbook, looping back to the first column when the end is reached.
class RollingChartDemo
{
    // Workbook components
    static Workbook workbook = null!;
    static Worksheet sheet = null!;
    static Chart chart = null!;

    // Parameters for the rolling window
    static int startColumn = 0;          // Index of the first column in the window (0 = A)
    static int windowSize = 10;          // Number of columns displayed in the chart
    static int totalColumns = 30;        // Total number of data columns available

    static void Main()
    {
        try
        {
            // Create a new workbook and fill it with sample data
            workbook = new Workbook();
            sheet = workbook.Worksheets[0];
            FillSampleData();

            // Add a line chart to the worksheet (Charts.Add returns the chart index)
            int chartIndex = sheet.Charts.Add(ChartType.Line, 15, 0, 30, 10);
            chart = sheet.Charts[chartIndex];

            UpdateChartSeries();

            // Save the initial workbook
            workbook.Save("RollingChart.xlsx");

            // Set up a timer to shift the data window every minute (60,000 ms)
            System.Timers.Timer timer = new System.Timers.Timer(60000);
            timer.Elapsed += OnTimerElapsed;
            timer.AutoReset = true;
            timer.Start();

            Console.WriteLine("Rolling chart started. Press Enter to stop.");
            Console.ReadLine();

            timer.Stop();
            timer.Dispose();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Populate the worksheet with random numbers for demonstration
    static void FillSampleData()
    {
        Random rnd = new Random();
        for (int col = 0; col < totalColumns; col++)
        {
            for (int row = 0; row < windowSize; row++)
            {
                sheet.Cells[row, col].PutValue(rnd.Next(0, 100));
            }
        }
    }

    // Re‑create the chart series to reflect the current window
    static void UpdateChartSeries()
    {
        // Clear any existing series
        chart.NSeries.Clear();

        // Determine the range for the current window
        int endColumn = startColumn + windowSize - 1;
        string startCell = CellsHelper.CellIndexToName(startColumn, 0);                 // e.g., A1
        string endCell = CellsHelper.CellIndexToName(endColumn, windowSize - 1);        // e.g., J10
        string range = $"{sheet.Name}!{startCell}:{endCell}";

        // Add the series using the calculated range
        chart.NSeries.Add(range, true);
        chart.Title.Text = $"Rolling Window: {startCell} – {endCell}";
    }

    // Timer callback: shift the window one column to the right
    static void OnTimerElapsed(object sender, ElapsedEventArgs e)
    {
        try
        {
            startColumn++;
            if (startColumn + windowSize > totalColumns)
            {
                // Loop back to the beginning when the end is reached
                startColumn = 0;
            }

            UpdateChartSeries();

            // Save the workbook so the updated chart is persisted
            workbook.Save("RollingChart.xlsx");
            Console.WriteLine($"Chart updated at {DateTime.Now}, window starts at column {CellsHelper.ColumnIndexToName(startColumn)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Timer error: {ex.Message}");
        }
    }
}
