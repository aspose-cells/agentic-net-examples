// Title: Measure Aspose.Cells Workbook Load Memory with KeepUnparsedData On vs Off (C#)
// Description: A C# console app that loads the same Excel file twice using Aspose.Cells—first with LoadOptions.KeepUnparsedData set to true (cache parsing enabled) and then with it set to false. It forces garbage collection before and after each load, records total memory, computes the difference, disposes the workbook, and prints the memory used for each scenario.
// Keywords: Aspose.Cells | C# | LoadOptions | KeepUnparsedData | memory benchmark | Excel workbook loading | cache parsing | memory footprint | GC.GetTotalMemory | performance testing
// Common Searches: Aspose.Cells memory usage KeepUnparsedData | disable cache parsing Aspose.Cells | measure memory consumption loading Excel with Aspose | C# Aspose.Cells memory benchmark | reduce RAM usage Aspose.Cells workbook
// Developer Intent: Determine how the KeepUnparsedData flag influences RAM consumption when loading a workbook with Aspose.Cells.
// Use Cases: Benchmark memory requirements of large workbooks to decide whether to retain unparsed data. | Lower server‑side Excel processing memory by disabling cache parsing. | Add automated tests that verify a measurable memory reduction when KeepUnparsedData is false.
// AI Prompts: Create a reusable C# method that returns the memory delta between loading a workbook with KeepUnparsedData true and false using Aspose.Cells. | Explain the internal data structures affected by the KeepUnparsedData option and why it impacts memory usage. | Provide guidelines for accurately measuring Aspose.Cells workbook load memory in .NET, including GC best practices.

using System;
using Aspose.Cells;

namespace AsposeCellsMemoryComparison
{
    // A C# console app that loads the same Excel file twice using Aspose.Cells—first with LoadOptions.KeepUnparsedData set to true (cache parsing enabled) and then with it set to false. It forces garbage collection before and after each load, records total memory, computes the difference, disposes the workbook, and prints the memory used for each scenario.
    class Program
    {
        static void Main()
        {
            // Path to the workbook that will be used for both loads.
            string filePath = "sample.xlsx";

            // ------------------------------------------------------------
            // Load with cache parsing (KeepUnparsedData) enabled (default).
            // ------------------------------------------------------------
            // Force a full garbage collection and get the baseline memory.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryBeforeEnabled = GC.GetTotalMemory(true);

            // Create LoadOptions with KeepUnparsedData = true.
            LoadOptions optionsEnabled = new LoadOptions();
            optionsEnabled.KeepUnparsedData = true; // cache parsing enabled

            // Load the workbook using the options.
            Workbook workbookEnabled = new Workbook(filePath, optionsEnabled);

            // Measure memory after loading.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryAfterEnabled = GC.GetTotalMemory(true);
            long memoryUsedEnabled = memoryAfterEnabled - memoryBeforeEnabled;

            // Release resources.
            workbookEnabled.Dispose();

            // ------------------------------------------------------------
            // Load with cache parsing disabled (KeepUnparsedData = false).
            // ------------------------------------------------------------
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryBeforeDisabled = GC.GetTotalMemory(true);

            // Create LoadOptions with KeepUnparsedData = false.
            LoadOptions optionsDisabled = new LoadOptions();
            optionsDisabled.KeepUnparsedData = false; // cache parsing disabled

            // Load the workbook using the options.
            Workbook workbookDisabled = new Workbook(filePath, optionsDisabled);

            // Measure memory after loading.
            GC.Collect();
            GC.WaitForPendingFinalizers();
            long memoryAfterDisabled = GC.GetTotalMemory(true);
            long memoryUsedDisabled = memoryAfterDisabled - memoryBeforeDisabled;

            // Release resources.
            workbookDisabled.Dispose();

            // ------------------------------------------------------------
            // Output the comparison results.
            // ------------------------------------------------------------
            Console.WriteLine($"Memory used with KeepUnparsedData = true : {memoryUsedEnabled:N0} bytes");
            Console.WriteLine($"Memory used with KeepUnparsedData = false: {memoryUsedDisabled:N0} bytes");
        }
    }
}
