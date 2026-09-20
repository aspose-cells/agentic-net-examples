// Title: Load a large Excel workbook with Aspose.Cells in C# and measure process private memory before and after calculation
// AI Prompts: Write C# code that opens a given .xlsx file using Aspose.Cells, executes CalculateFormula, and prints the process's PrivateMemorySize64 value before and after the operation. | Show how to capture and display the memory consumption of a .NET process while loading and processing a large workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# example to monitor memory usage when opening a large .xlsx file | how to get process private memory before and after CalculateFormula in .NET | measure memory delta of Aspose.Cells workbook load in C# | C# code to track memory consumption while processing large Excel workbooks with Aspose.Cells | profile memory usage of Aspose.Cells operations in a .NET application
// Tags: Aspose.Cells workbook loading memory profiling | Aspose.Cells process memory tracking | C# calculate formula memory measurement | large Excel file handling Aspose.Cells | memory delta analysis .NET process

using System;
using System.Diagnostics;
using System.IO;
using Aspose.Cells;

// The program verifies the existence of a large Excel file, records the current process's private memory size, loads the workbook with Aspose.Cells, optionally calculates all formulas, records the memory size again, and outputs the before, after, and delta memory values, all wrapped in error handling.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the workbook
            string inputPath = "LargeWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Record process memory before loading the workbook
            Process currentProcess = Process.GetCurrentProcess();
            long memoryBefore = currentProcess.PrivateMemorySize64;

            // Load the workbook (standard Aspose.Cells API)
            Workbook workbook = new Workbook(inputPath);

            // Optionally perform an operation to ensure the workbook is fully processed
            // For example, calculate all formulas
            workbook.CalculateFormula();

            // Record process memory after loading and processing the workbook
            long memoryAfter = currentProcess.PrivateMemorySize64;

            // Output overall process memory change
            Console.WriteLine("=== Process Memory Change ===");
            Console.WriteLine($"Memory before processing (bytes): {memoryBefore}");
            Console.WriteLine($"Memory after processing (bytes): {memoryAfter}");
            Console.WriteLine($"Memory delta (bytes): {memoryAfter - memoryBefore}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
