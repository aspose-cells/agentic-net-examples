// Title: How to time Worksheet.FreezePanes calls with Stopwatch and log the results in Aspose.Cells for .NET
// Common Searches: C# benchmark Worksheet.FreezePanes performance using Stopwatch in Aspose.Cells | log execution time of FreezePanes method in Aspose.Cells .NET | measure time taken by each FreezePanes call in a workbook with Aspose.Cells | record duration of FreezePanes operation for different rows and columns in C# | Aspose.Cells example for measuring FreezePanes latency
// Tags: benchmark worksheet freeze panes Aspose.Cells | stopwatch timing Aspose.Cells FreezePanes | measure freeze panes latency C# | log FreezePanes execution time .NET | performance monitoring Aspose.Cells worksheet

using Aspose.Cells;
using System;
using System.Diagnostics;

// The sample creates a Workbook, calls a helper method three times with different row and column indices, and each helper uses Stopwatch.StartNew() to time the sheet.FreezePanes(row, column, row, column) call. After stopping the timer, it writes the elapsed milliseconds to the console, then saves the workbook as FreezePanesDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.Worksheets[0];

            // Record duration of each FreezePanes call
            RecordFreezePanes(sheet, 1, 1);
            RecordFreezePanes(sheet, 2, 2);
            RecordFreezePanes(sheet, 3, 3);

            // Save the workbook
            workbook.Save("FreezePanesDemo.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Measures and logs the time taken to apply FreezePanes
    static void RecordFreezePanes(Worksheet sheet, int row, int column)
    {
        try
        {
            Stopwatch sw = Stopwatch.StartNew();

            // Freeze rows above 'row' and columns left of 'column'
            // The overload requires total rows and columns to freeze; using the same values freezes up to the specified cell.
            sheet.FreezePanes(row, column, row, column);

            sw.Stop();

            Console.WriteLine($"FreezePanes(row:{row}, column:{column}) took {sw.Elapsed.TotalMilliseconds} ms");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"FreezePanes failed for row:{row}, column:{column}. Error: {ex.Message}");
        }
    }
}
