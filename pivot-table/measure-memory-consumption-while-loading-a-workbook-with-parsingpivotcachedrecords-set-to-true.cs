// Title: Measure memory consumption of loading a workbook with ParsingPivotCachedRecords = true in Aspose.Cells for .NET
// Description: C# example that enables LoadOptions.ParsingPivotCachedRecords, captures Process.PrivateMemorySize64 before and after creating a Workbook, computes the memory delta, and outputs the bytes used. Demonstrates how to benchmark the memory impact of pivot‑cache parsing in Aspose.Cells.
// Keywords: Aspose.Cells memory benchmark | ParsingPivotCachedRecords performance | C# load workbook memory usage | Aspose.Cells LoadOptions | private memory size .NET | Excel pivot cache memory | measure workbook load overhead | Aspose.Cells .NET performance
// Common Searches: how to measure memory usage when loading an Excel file with Aspose.Cells | ParsingPivotCachedRecords memory impact Aspose.Cells | C# Aspose.Cells load options memory consumption | benchmark private memory before and after Workbook creation | performance of pivot cache parsing in Aspose.Cells
// Developer Intent: Determine the exact amount of memory allocated when a workbook is loaded with the ParsingPivotCachedRecords option enabled.
// Use Cases: Benchmark memory overhead of pivot‑cache parsing for large Excel files. | Validate that a server meets memory requirements before processing workbooks with cached pivot data. | Compare memory footprints of default loading versus loading with ParsingPivotCachedRecords set to true.
// AI Prompts: Generate a C# snippet that logs memory usage at several stages while loading a workbook with ParsingPivotCachedRecords enabled using Aspose.Cells. | Explain how to interpret differences in Process.PrivateMemorySize64 to assess the memory cost of pivot cache parsing. | Suggest best practices to minimize memory consumption when loading workbooks with ParsingPivotCachedRecords set to true.

using System;
using System.Diagnostics;
using Aspose.Cells;

namespace AsposeCellsMemoryMeasurement
{
    // C# example that enables LoadOptions.ParsingPivotCachedRecords, captures Process.PrivateMemorySize64 before and after creating a Workbook, computes the memory delta, and outputs the bytes used. Demonstrates how to benchmark the memory impact of pivot‑cache parsing in Aspose.Cells.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the Excel file to be loaded
            string filePath = "sample.xlsx";

            // Create load options and enable parsing of pivot cached records
            LoadOptions loadOptions = new LoadOptions();
            loadOptions.ParsingPivotCachedRecords = true;

            // Measure memory before loading
            long memoryBefore = Process.GetCurrentProcess().PrivateMemorySize64;

            // Load the workbook with the specified options
            Workbook workbook = new Workbook(filePath, loadOptions);

            // Measure memory after loading
            long memoryAfter = Process.GetCurrentProcess().PrivateMemorySize64;

            // Calculate and display the memory consumption
            long memoryUsed = memoryAfter - memoryBefore;
            Console.WriteLine($"Memory used for loading workbook with ParsingPivotCachedRecords=true: {memoryUsed} bytes");

            // Optional: keep the workbook alive for further processing
            // ...

            // Dispose the workbook if no longer needed
            workbook.Dispose();
        }
    }
}
