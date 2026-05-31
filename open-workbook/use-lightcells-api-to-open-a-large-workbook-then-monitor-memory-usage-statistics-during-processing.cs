using System;
using System.Diagnostics;
using Aspose.Cells;

public class LightCellsMemoryMonitor : LightCellsDataHandler
{
    // Called when a worksheet starts processing
    public bool StartSheet(Worksheet sheet)
    {
        Console.WriteLine($"Start processing sheet: {sheet.Name}");
        PrintMemory();
        return true; // Continue processing this sheet
    }

    // Called before a row is processed
    public bool StartRow(int rowIndex)
    {
        Console.WriteLine($"Start processing row: {rowIndex}");
        PrintMemory();
        return true; // Continue processing this row
    }

    // Called after a row is read (optional processing)
    public bool ProcessRow(Row row)
    {
        // No custom row logic needed for this demo
        return true; // Continue processing
    }

    // Called before a cell is processed
    public bool StartCell(int columnIndex)
    {
        // No custom cell start logic needed
        return true; // Continue processing this cell
    }

    // Called after a cell is read (optional processing)
    public bool ProcessCell(Cell cell)
    {
        // No custom cell logic needed for this demo
        return true; // Continue processing
    }

    // Helper to output current private memory usage in MB
    private void PrintMemory()
    {
        long bytes = Process.GetCurrentProcess().PrivateMemorySize64;
        Console.WriteLine($"Current private memory: {bytes / 1024 / 1024} MB");
    }
}

public class Program
{
    public static void Main()
    {
        // Path to the large workbook to be processed
        string inputPath = "LargeWorkbook.xlsx";

        // Create LoadOptions and assign the custom LightCellsDataHandler
        LoadOptions loadOptions = new LoadOptions();
        loadOptions.LightCellsDataHandler = new LightCellsMemoryMonitor();

        // Optional: set memory mode to reduce memory consumption while loading
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the workbook using LightCells mode (streaming, low memory)
        Workbook workbook = new Workbook(inputPath, loadOptions);

        // After loading completes, display final memory usage
        Console.WriteLine("Finished loading workbook.");
        long finalBytes = Process.GetCurrentProcess().PrivateMemorySize64;
        Console.WriteLine($"Final private memory: {finalBytes / 1024 / 1024} MB");
    }
}