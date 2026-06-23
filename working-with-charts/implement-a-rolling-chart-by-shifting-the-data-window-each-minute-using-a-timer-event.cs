using System;
using System.IO;
using System.Timers;
using Aspose.Cells;
using Aspose.Cells.Charts;

class RollingChartDemo
{
    // Workbook and worksheet that hold the data and the chart
    static Workbook workbook;
    static Worksheet sheet;
    static Chart chart;

    // Size of the visible window on the chart
    const int windowSize = 10;

    // Row where the current window starts (Excel rows are 1‑based)
    static int startRow = 2; // data begins at row 2 (A2, B2)

    // Total number of data rows available
    const int totalRows = 30;

    // Timer that triggers the shift every minute
    static System.Timers.Timer timer;

    static void Main()
    {
        try
        {
            // ---------- Create workbook and fill sample time‑series data ----------
            workbook = new Workbook();
            sheet = workbook.Worksheets[0];

            sheet.Cells["A1"].PutValue("Date");
            sheet.Cells["B1"].PutValue("Value");

            DateTime baseDate = DateTime.Today;
            for (int i = 0; i < totalRows; i++)
            {
                // Column A – dates, Column B – numeric values
                sheet.Cells[1 + i, 0].PutValue(baseDate.AddDays(i));
                sheet.Cells[1 + i, 1].PutValue(i * 10 + 5);
            }

            // ---------- Add a line chart that initially shows the first window ----------
            int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
            chart = sheet.Charts[chartIndex];
            UpdateChartRange();               // set initial data range
            chart.Calculate();                // ensure the chart is rendered

            // ---------- Configure timer to shift the window each minute ----------
            timer = new System.Timers.Timer(60 * 1000); // 60 000 ms = 1 minute
            timer.Elapsed += OnTimerElapsed;
            timer.AutoReset = true;
            timer.Start();

            Console.WriteLine("Rolling chart is running. Press ENTER to stop.");
            Console.ReadLine();               // keep the application alive

            timer.Stop();                     // stop timer when user ends the program
            timer.Dispose();

            // Save final workbook
            string finalPath = "RollingChart_Final.xlsx";
            workbook.Save(finalPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(finalPath)}");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }

    // Timer callback – moves the window forward by one row and updates the chart
    static void OnTimerElapsed(object sender, ElapsedEventArgs e)
    {
        try
        {
            // Advance start row; wrap to the beginning when we reach the end
            startRow++;
            if (startRow + windowSize - 1 > totalRows + 1) // +1 because data starts at row 2
            {
                startRow = 2;
            }

            UpdateChartRange();   // apply new data range to the chart
            chart.Calculate();    // recalculate chart to reflect changes

            // Optional: save a snapshot after each shift (useful for debugging)
            string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string snapshotPath = $"RollingChart_{timestamp}.xlsx";

            // Ensure we don't overwrite an existing file unintentionally
            if (!File.Exists(snapshotPath))
            {
                workbook.Save(snapshotPath);
            }
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"Timer error: {ex.Message}");
        }
    }

    // Helper that builds the range strings for the current window and assigns them to the chart
    static void UpdateChartRange()
    {
        string categoryRange = $"A{startRow}:A{startRow + windowSize - 1}";
        string valueRange    = $"B{startRow}:B{startRow + windowSize - 1}";

        // Clear any existing series and create a fresh one with the new ranges
        chart.NSeries.Clear();
        chart.NSeries.Add(valueRange, true);
        chart.NSeries.CategoryData = categoryRange;
    }
}