// Title: Measure and log process memory before and after converting an Excel workbook to multi‑page TIFF with Aspose.Cells in C#
// AI Prompts: Write C# code that captures the current process's private memory size, forces garbage collection, loads an .xlsx file using Aspose.Cells, saves it as a multi‑page TIFF, and prints the memory usage before and after the conversion. | Create a C# snippet that benchmarks memory consumption of Aspose.Cells when exporting a workbook to TIFF, including explicit GC calls and calculating the memory delta.
// Common Searches: how to check memory usage in C# when saving Excel as TIFF with Aspose.Cells | C# profiling private memory before and after Aspose.Cells TIFF export | measure process memory impact of converting .xlsx to .tiff using Aspose.Cells | log memory consumption of Aspose.Cells workbook.Save to TIFF in .NET
// Tags: Aspose.Cells TIFF export memory usage | C# process private memory measurement | force garbage collection for accurate memory reading | Excel workbook to multi-page TIFF performance | log memory delta after Aspose.Cells conversion

using System;
using System.Diagnostics;
using Aspose.Cells;

// // This example measures the private memory of the current process before and after loading an Excel workbook and saving it as a multi‑page TIFF with Aspose.Cells, using forced garbage collection to obtain accurate readings and outputting the memory delta.
class Program
{
    static void Main()
    {
        // Paths (adjust as needed)
        string inputFile = "input.xlsx";
        string outputFile = "output.tiff";

        // Ensure Aspose.Cells license is set if you have one
        // License license = new License();
        // license.SetLicense("Aspose.Cells.NET.lic");

        // Get current process for memory measurement
        Process proc = Process.GetCurrentProcess();

        // Force garbage collection before measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Memory before conversion (in MB)
        long memoryBefore = proc.PrivateMemorySize64 / (1024 * 1024);
        Console.WriteLine($"Memory before TIFF conversion: {memoryBefore} MB");

        // Load the workbook
        Workbook workbook = new Workbook(inputFile);

        // Save as TIFF (each worksheet will be saved as a separate page)
        workbook.Save(outputFile, SaveFormat.Tiff);

        // Force garbage collection again before second measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Memory after conversion (in MB)
        long memoryAfter = proc.PrivateMemorySize64 / (1024 * 1024);
        Console.WriteLine($"Memory after TIFF conversion: {memoryAfter} MB");

        // Optional: display the difference
        Console.WriteLine($"Memory change: {memoryAfter - memoryBefore} MB");
    }
}
