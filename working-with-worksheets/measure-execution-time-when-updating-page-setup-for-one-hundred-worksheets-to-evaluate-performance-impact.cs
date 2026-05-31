using System;
using System.Diagnostics;
using Aspose.Cells;

class MeasurePageSetupPerformance
{
    static void Main()
    {
        // Create a new workbook and ensure it has 100 worksheets
        Workbook workbook = new Workbook();
        while (workbook.Worksheets.Count < 100)
        {
            workbook.Worksheets.Add();
        }

        // Start measuring time
        Stopwatch stopwatch = Stopwatch.StartNew();

        // Update page setup for each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            PageSetup setup = sheet.PageSetup;
            setup.FitToPagesWide = 1;   // fit columns to one page width
            setup.FitToPagesTall = 0;   // let height adjust automatically
            setup.Orientation = PageOrientationType.Landscape;
        }

        // Stop measuring
        stopwatch.Stop();

        Console.WriteLine($"Updating page setup for 100 worksheets took {stopwatch.ElapsedMilliseconds} ms.");

        // Save the workbook (optional, demonstrates lifecycle usage)
        workbook.Save("PageSetupPerformance.xlsx");
    }
}