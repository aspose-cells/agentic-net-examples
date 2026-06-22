using System;
using System.Diagnostics;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Measure first FreezePanes call (by indices)
        Stopwatch stopwatch = Stopwatch.StartNew();
        worksheet.FreezePanes(3, 3, 3, 3);
        stopwatch.Stop();
        Console.WriteLine($"FreezePanes(3,3,3,3) elapsed: {stopwatch.ElapsedMilliseconds} ms");

        // Measure second FreezePanes call (by cell name)
        stopwatch.Restart();
        worksheet.FreezePanes("E5", 5, 5);
        stopwatch.Stop();
        Console.WriteLine($"FreezePanes(\"E5\",5,5) elapsed: {stopwatch.ElapsedMilliseconds} ms");

        // Save the workbook
        workbook.Save("FreezePanesTiming.xlsx");
    }
}