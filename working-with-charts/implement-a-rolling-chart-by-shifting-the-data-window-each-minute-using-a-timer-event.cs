// Title: C# Rolling Line Chart that Shifts Data Window Every Minute with Aspose.Cells
// Description: Creates an Excel workbook, fills it with dates and random values, adds a line chart with a time‑scaled X‑axis, and uses System.Timers.Timer to move a fixed‑size data window down one row each minute. The chart series range is updated and recalculated automatically before the file is saved.
// Keywords: Aspose.Cells | C# | .NET | rolling chart | dynamic chart range | timer update | real‑time Excel chart | time‑scaled axis | line chart programmatically | Excel automation | chart refresh
// Common Searches: Aspose.Cells rolling chart example | C# update Excel chart every minute | timer based chart data window Aspose.Cells | dynamic line chart with time axis .NET | how to shift chart series range programmatically
// Developer Intent: The developer wants to generate a line chart that automatically scrolls through a predefined number of rows, advancing the window at one‑minute intervals.
// Use Cases: Live sensor or IoT data visualization that scrolls in near‑real time. | Continuous sales or KPI dashboard where the latest values replace the oldest. | Animated presentation of historical trends that moves forward automatically.
// AI Prompts: Show how to replace System.Timers.Timer with an async/await loop for chart updates. | Provide code to export the chart as a PNG after each timer tick. | Explain modifications needed to use a 15‑row window and a 30‑second interval.

using System;
using System.IO;
using System.Timers;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Creates an Excel workbook, fills it with dates and random values, adds a line chart with a time‑scaled X‑axis, and uses System.Timers.Timer to move a fixed‑size data window down one row each minute. The chart series range is updated and recalculated automatically before the file is saved.
class RollingChartDemo
{
    // Configuration
    const int TotalRows = 30;          // Total data rows (including header)
    const int WindowSize = 10;         // Number of rows displayed in the chart window
    const int StartDataRow = 2;        // First data row (1‑based index in Excel)
    const int DateColumn = 0;          // Column A
    const int ValueColumn = 1;         // Column B
    const int ChartTopRow = 5;
    const int ChartLeftColumn = 0;
    const int ChartBottomRow = 20;
    const int ChartRightColumn = 8;
    const double TimerIntervalMs = 60_000; // 1 minute

    static Workbook workbook;
    static Worksheet sheet;
    static Chart chart;
    static int currentStartRow = StartDataRow; // Tracks the first row of the current window

    static void Main()
    {
        try
        {
            // ---------- Create workbook and populate sample data ----------
            workbook = new Workbook();
            sheet = workbook.Worksheets[0];

            // Header
            sheet.Cells[0, DateColumn].PutValue("Date");
            sheet.Cells[0, ValueColumn].PutValue("Value");

            // Populate dates (today + i days) and random values
            Random rnd = new Random();
            for (int i = 0; i < TotalRows - 1; i++)
            {
                int row = i + 1; // Excel rows are 0‑based in Aspose.Cells
                sheet.Cells[row, DateColumn].PutValue(DateTime.Today.AddDays(i));
                sheet.Cells[row, ValueColumn].PutValue(rnd.Next(50, 150));
            }

            // ---------- Add a line chart ----------
            int chartIndex = sheet.Charts.Add(ChartType.Line, ChartTopRow, ChartLeftColumn, ChartBottomRow, ChartRightColumn);
            chart = sheet.Charts[chartIndex];

            // Initial data window
            UpdateChartDataRange();

            // Optional: format axes as time scale
            chart.CategoryAxis.CategoryType = CategoryType.TimeScale;
            chart.CategoryAxis.MinorUnitScale = TimeUnit.Days;
            chart.CategoryAxis.MinorUnit = 1;

            // ---------- Set up a timer to shift the window ----------
            System.Timers.Timer timer = new System.Timers.Timer(TimerIntervalMs);
            timer.Elapsed += OnTimerElapsed;
            timer.AutoReset = true;
            timer.Start();

            Console.WriteLine("Rolling chart started. Press ENTER to stop...");
            Console.ReadLine();

            timer.Stop();
            timer.Dispose();

            // Save the workbook (you can open it in Excel to see the final state)
            string outputPath = "RollingChartDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Timer callback: shift the window by one row and refresh the chart
    private static void OnTimerElapsed(object? sender, ElapsedEventArgs e)
    {
        try
        {
            // Move start row down, wrap around when reaching the end of data
            if (currentStartRow + WindowSize > TotalRows)
                currentStartRow = StartDataRow;
            else
                currentStartRow++;

            UpdateChartDataRange();

            // Recalculate the chart so the visual reflects the new range
            chart.Calculate();

            Console.WriteLine($"Chart window updated: rows {currentStartRow} to {currentStartRow + WindowSize - 1} at {DateTime.Now}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Timer error: {ex.Message}");
        }
    }

    // Helper: builds the range string for the current window and applies it to the chart
    private static void UpdateChartDataRange()
    {
        // Build range strings like "B2:B11" for values and "A2:A11" for categories
        string valueRange = $"B{currentStartRow + 1}:B{currentStartRow + WindowSize}";
        string categoryRange = $"A{currentStartRow + 1}:A{currentStartRow + WindowSize}";

        // Clear existing series and add a new one with values
        chart.NSeries.Clear();
        chart.NSeries.Add(valueRange, true);
        // Assign category (X‑axis) data using XValues (compatible with all Aspose.Cells versions)
        chart.NSeries[0].XValues = categoryRange;
    }
}
