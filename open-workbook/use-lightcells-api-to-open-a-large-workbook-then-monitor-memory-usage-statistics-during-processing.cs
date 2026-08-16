// Title: C# – Monitor memory while loading a large workbook with Aspose.Cells LightCells API
// Description: A .NET example that creates a custom LightCellsDataHandler to log the process's private memory (MB) at the start of each worksheet and row. The workbook is opened in LightCells mode with MemorySetting.MemoryPreference, then saved, demonstrating low‑memory processing of huge Excel files.
// Keywords: Aspose.Cells | LightCells | memory monitoring | C# | .NET | large Excel workbook | MemorySetting.MemoryPreference | LoadOptions | custom LightCellsDataHandler | Process.GetCurrentProcess | PrivateMemorySize64 | performance optimization | GitHub example | Excel file loading
// Common Searches: Aspose.Cells LightCells memory logging example | C# track memory usage when loading big Excel file | How to use MemorySetting.MemoryPreference with LightCells | Custom LightCellsDataHandler to monitor RAM consumption | Load large workbook in .NET with low memory footprint
// Developer Intent: I need to load a massive Excel workbook using LightCells and capture memory consumption at key processing points.
// Use Cases: Identify memory spikes during sheet‑level processing of a huge workbook. | Diagnose rows that cause unexpected RAM growth while streaming data. | Combine MemoryPreference mode with a custom handler to keep the footprint under control for server‑side batch jobs.
// AI Prompts: Generate a LightCellsDataHandler that writes memory metrics to a CSV file instead of the console. | Show how to add a progress bar that updates together with memory logs during LightCells loading. | Explain trade‑offs between MemoryPreference, MemorySetting.Default, and MemorySetting.Performance for very large workbooks.

using System;
using System.Diagnostics;
using Aspose.Cells;

// A .NET example that creates a custom LightCellsDataHandler to log the process's private memory (MB) at the start of each worksheet and row. The workbook is opened in LightCells mode with MemorySetting.MemoryPreference, then saved, demonstrating low‑memory processing of huge Excel files.
class MemoryMonitoringHandler : LightCellsDataHandler
{
    // Called when a worksheet starts processing
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Start processing sheet: {sheet.Name}");
        LogMemory("StartSheet");
        return true; // Continue processing this sheet
    }

    // Called before a row is processed
    public bool StartRow(int rowIndex)
    {
        Console.WriteLine($"Start processing row: {rowIndex}");
        LogMemory($"StartRow {rowIndex}");
        return true; // Continue processing this row
    }

    // Optional: process row data (not used here)
    public bool ProcessRow(Row row) => true;

    // Optional: called before a cell is processed
    public bool StartCell(int columnIndex) => true;

    // Optional: process cell data (not used here)
    public bool ProcessCell(Cell cell) => true;

    // Helper to output current memory usage in MB
    private void LogMemory(string stage)
    {
        long memoryMb = Process.GetCurrentProcess().PrivateMemorySize64 / (1024 * 1024);
        Console.WriteLine($"{stage} - Memory usage: {memoryMb} MB");
    }
}

class Program
{
    static void Main()
    {
        // Paths to the input large workbook and the output file
        string inputPath = "LargeWorkbook.xlsx";
        string outputPath = "ProcessedWorkbook.xlsx";

        // Create load options and assign the custom LightCellsDataHandler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new MemoryMonitoringHandler();

        // Optionally set a memory‑optimized mode for loading
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the workbook using LightCells mode (uses the provided rule)
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // After loading, you can perform additional operations if needed.
        // For demonstration, simply save the workbook.
        workbook.Save(outputPath);

        Console.WriteLine("Workbook processing completed.");
    }
}
