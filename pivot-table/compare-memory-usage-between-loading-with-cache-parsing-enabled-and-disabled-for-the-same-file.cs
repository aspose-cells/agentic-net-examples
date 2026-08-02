using System;
using System.Diagnostics;
using Aspose.Cells;

class CompareMemoryUsage
{
    static void Main()
    {
        // Path to the workbook to be loaded for comparison
        string filePath = "sample.xlsx";

        // -------------------------------------------------
        // Load with cache parsing (KeepUnparsedData = true)
        // -------------------------------------------------
        LoadOptions optionsWithCache = new LoadOptions();
        optionsWithCache.KeepUnparsedData = true; // cache parsing enabled

        // Ensure a clean memory state before measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();

        long memBeforeCache = GC.GetTotalMemory(true);
        Workbook wbWithCache = new Workbook(filePath, optionsWithCache);
        long memAfterCache = GC.GetTotalMemory(true);
        long usedWithCache = memAfterCache - memBeforeCache;

        // -------------------------------------------------
        // Load without cache parsing (KeepUnparsedData = false)
        // -------------------------------------------------
        LoadOptions optionsWithoutCache = new LoadOptions();
        optionsWithoutCache.KeepUnparsedData = false; // cache parsing disabled

        // Clean memory again before second measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();

        long memBeforeNoCache = GC.GetTotalMemory(true);
        Workbook wbWithoutCache = new Workbook(filePath, optionsWithoutCache);
        long memAfterNoCache = GC.GetTotalMemory(true);
        long usedWithoutCache = memAfterNoCache - memBeforeNoCache;

        // Output the memory usage comparison
        Console.WriteLine($"Memory used with KeepUnparsedData = true : {usedWithCache} bytes");
        Console.WriteLine($"Memory used with KeepUnparsedData = false: {usedWithoutCache} bytes");

        // Dispose workbooks if needed
        wbWithCache.Dispose();
        wbWithoutCache.Dispose();
    }
}