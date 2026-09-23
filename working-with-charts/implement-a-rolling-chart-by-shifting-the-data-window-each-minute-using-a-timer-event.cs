// Title: How to create a rolling line chart in Excel that updates every minute using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells to generate a line chart that shows a 10‑row moving window and refreshes the series each minute using System.Timers.Timer. | Adjust the sample to use a 30‑second timer interval and display the latest 20 rows of data in the chart. | Extend the rolling chart by adding a second series that computes and plots a moving average of the values.
// Common Searches: Aspose.Cells C# example for updating an Excel chart every minute | How to shift the data range of a chart with a timer in .NET | Create a dynamic line chart that scrolls with new data using Aspose.Cells | C# timer based rolling data window for Excel chart
// Tags: rolling line chart Aspose.Cells | timer driven chart series update C# | dynamic Excel chart data window .NET | shift chart series range programmatically | Aspose.Cells moving data range example

using System;
using System.IO;
using System.Timers;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace RollingChartDemo
{
    // // Demonstrates creating an Excel workbook with timestamp and random value data, adding a line chart, and using a System.Timers.Timer to shift a 10‑row data window each minute, updating the chart series and saving the workbook.
    class Program
    {
        // Path to the Excel file that will hold the chart.
        private const string WorkbookPath = "RollingChart.xlsx";

        // Number of rows to display in the rolling window.
        private const int WindowSize = 10;

        // Total number of data rows available.
        private const int TotalRows = 100;

        // Current start row of the window (1‑based, includes header row).
        private static int _currentStartRow = 2; // Assuming row 1 has headers.

        // Aspose.Cells objects that need to be accessed from the timer callback.
        private static Workbook _workbook = null!;
        private static Worksheet _dataSheet = null!;
        private static Chart _chart = null!;

        static void Main()
        {
            try
            {
                // -------------------------------------------------
                // 1. Create a new workbook and populate sample data.
                // -------------------------------------------------
                _workbook = new Workbook();
                _dataSheet = _workbook.Worksheets[0];
                _dataSheet.Name = "Data";

                // Header row.
                _dataSheet.Cells["A1"].PutValue("Timestamp");
                _dataSheet.Cells["B1"].PutValue("Value");

                // Fill sample data (e.g., timestamps at 1‑minute intervals and random values).
                DateTime startTime = DateTime.Now.AddMinutes(-TotalRows);
                Random rnd = new Random();
                for (int i = 0; i < TotalRows; i++)
                {
                    _dataSheet.Cells[i + 2, 0].PutValue(startTime.AddMinutes(i)); // Column A
                    _dataSheet.Cells[i + 2, 1].PutValue(rnd.NextDouble() * 100); // Column B
                }

                // -------------------------------------------------
                // 2. Create a line chart that will display the window.
                // -------------------------------------------------
                int chartIndex = _dataSheet.Charts.Add(ChartType.Line, 5, 0, 25, 10);
                _chart = _dataSheet.Charts[chartIndex];
                _chart.Title.Text = "Rolling Data Window";

                // Initial series using the first window.
                UpdateChartSeries();

                // -------------------------------------------------
                // 3. Save the initial workbook.
                // -------------------------------------------------
                EnsureDirectoryExists(WorkbookPath);
                _workbook.Save(WorkbookPath);

                // -------------------------------------------------
                // 4. Set up a timer to shift the window every minute.
                // -------------------------------------------------
                System.Timers.Timer timer = new System.Timers.Timer(60_000); // 60,000 ms = 1 minute
                timer.Elapsed += OnTimerElapsed;
                timer.AutoReset = true;
                timer.Start();

                Console.WriteLine("Rolling chart started. Press Enter to exit...");
                Console.ReadLine();

                // Clean up.
                timer.Stop();
                timer.Dispose();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // -------------------------------------------------
        // Timer callback: shift the data window and refresh the chart.
        // -------------------------------------------------
        private static void OnTimerElapsed(object? sender, ElapsedEventArgs e)
        {
            try
            {
                // Move the window one row down. Loop back to the start when reaching the end.
                _currentStartRow++;
                if (_currentStartRow + WindowSize - 1 > TotalRows + 1) // +1 because of header row
                {
                    _currentStartRow = 2; // Reset to first data row.
                }

                // Update the chart series to point to the new range.
                UpdateChartSeries();

                // Save the workbook so the changes are visible in Excel.
                EnsureDirectoryExists(WorkbookPath);
                _workbook.Save(WorkbookPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Timer error: {ex.Message}");
            }
        }

        // -------------------------------------------------
        // Helper: rebuild the chart series with the current window range.
        // -------------------------------------------------
        private static void UpdateChartSeries()
        {
            // Build the address strings for the category (X) and values (Y) ranges.
            string categoryRange = $"Data!$A${_currentStartRow}:$A${_currentStartRow + WindowSize - 1}";
            string valuesRange   = $"Data!$B${_currentStartRow}:$B${_currentStartRow + WindowSize - 1}";

            // Recreate the series for the current window.
            _chart.NSeries.Clear();
            _chart.NSeries.Add(valuesRange, true);
            // If the Aspose.Cells version supports CategoryData, uncomment the next line:
            // _chart.NSeries[0].CategoryData = categoryRange;
        }

        // Ensure the directory for the workbook exists to avoid FileNotFoundException.
        private static void EnsureDirectoryExists(string filePath)
        {
            try
            {
                string? directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Directory creation error: {ex.Message}");
                throw;
            }
        }
    }
}
