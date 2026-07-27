using System;
using System.Diagnostics;
using Aspose.Cells;

class WorkbookProcessingLogger
{
    static void Main()
    {
        // Stopwatch to measure each step
        Stopwatch sw = new Stopwatch();

        // -------------------- Create a new workbook --------------------
        sw.Start();
        Workbook newWb = new Workbook();               // creates an empty workbook
        sw.Stop();
        Console.WriteLine($"Workbook creation time: {sw.ElapsedMilliseconds} ms");
        sw.Reset();

        // -------------------- Load an existing workbook --------------------
        sw.Start();
        LoadOptions loadOpts = new LoadOptions();      // default load options
        Workbook wb = new Workbook("input.xlsx", loadOpts);
        sw.Stop();
        Console.WriteLine($"Workbook loading time: {sw.ElapsedMilliseconds} ms");
        sw.Reset();

        // -------------------- Calculate all formulas --------------------
        sw.Start();
        wb.CalculateFormula();                        // performs full formula calculation
        sw.Stop();
        Console.WriteLine($"Formula calculation time: {sw.ElapsedMilliseconds} ms");
        sw.Reset();

        // -------------------- Save the workbook --------------------
        sw.Start();
        wb.Save("output.xlsx");                        // saves the workbook to disk
        sw.Stop();
        Console.WriteLine($"Workbook saving time: {sw.ElapsedMilliseconds} ms");
    }
}