// Title: Compare memory usage of Automatic and AutomaticExceptTables load modes for large XLSX workbooks with Aspose.Cells for .NET
// Common Searches: how to benchmark Aspose.Cells MemorySetting Automatic versus AutomaticExceptTables in C# | measure memory consumption when loading a 50,000 row Excel file with Aspose.Cells LoadOptions | C# code to compare memory usage of different LoadOptions.MemorySetting values | profile .NET memory while loading large XLSX workbooks using Aspose.Cells | performance test for Aspose.Cells workbook loading memory impact
// Tags: Aspose.Cells LoadOptions MemorySetting comparison | automatic memory setting for large XLSX load | automaticexcepttables load mode memory usage | large workbook memory profiling .NET | benchmark Aspose.Cells workbook loading performance

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a 50,000‑row by 50‑column XLSX workbook, saves it to a temporary file, then loads the file twice using Aspose.Cells LoadOptions with the MemorySetting property set to "Automatic" and "AutomaticExceptTables" (applied via reflection). It records GC memory before and after each load, computes the memory delta, outputs the usage in megabytes, and finally deletes the temporary workbook.
class MemoryComparison
{
    static void Main()
    {
        // Path for the temporary large workbook
        string filePath = Path.Combine(Path.GetTempPath(), "LargeWorkbook.xlsx");

        // Create a large workbook with sample data
        CreateLargeWorkbook(filePath);

        // Measure memory consumption when loading with Automatic mode (if supported)
        long memoryAuto = MeasureMemoryLoad(filePath, "Auto");
        Console.WriteLine($"Memory used with Automatic mode: {memoryAuto / 1024.0 / 1024.0:F2} MB");

        // Measure memory consumption when loading with AutomaticExceptTables mode (if supported)
        long memoryAutoExceptTables = MeasureMemoryLoad(filePath, "AutoExceptTables");
        Console.WriteLine($"Memory used with AutomaticExceptTables mode: {memoryAutoExceptTables / 1024.0 / 1024.0:F2} MB");

        // Clean up the temporary file
        try
        {
            if (File.Exists(filePath))
                File.Delete(filePath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to delete temporary file: {ex.Message}");
        }
    }

    // Creates a workbook with a large number of rows and columns filled with data
    static void CreateLargeWorkbook(string path)
    {
        const int totalRows = 50000;
        const int totalColumns = 50;

        var wb = new Workbook();
        var ws = wb.Worksheets[0];

        // Populate cells with sample numeric data
        for (int row = 0; row < totalRows; row++)
        {
            for (int col = 0; col < totalColumns; col++)
            {
                ws.Cells[row, col].PutValue(row * col);
            }
        }

        // Save the workbook to the specified path
        wb.Save(path);
        wb.Dispose();
    }

    // Loads the workbook using the specified memory setting name (if supported) and returns the memory delta
    static long MeasureMemoryLoad(string path, string settingName)
    {
        if (!File.Exists(path))
        {
            Console.WriteLine($"File not found: {path}");
            return 0;
        }

        // Force garbage collection before measurement
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();

        // Record memory usage before loading
        long beforeLoad = GC.GetTotalMemory(true);

        // Prepare load options
        var loadOptions = new LoadOptions(LoadFormat.Xlsx);

        // Attempt to set the MemorySetting property via reflection (if the enum/value exists)
        try
        {
            var prop = typeof(LoadOptions).GetProperty("MemorySetting");
            if (prop != null && prop.PropertyType.IsEnum)
            {
                var enumType = prop.PropertyType;
                if (Enum.IsDefined(enumType, settingName))
                {
                    var enumValue = Enum.Parse(enumType, settingName, ignoreCase: true);
                    prop.SetValue(loadOptions, enumValue);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unable to apply memory setting '{settingName}': {ex.Message}");
        }

        // Load the workbook
        Workbook wb = null;
        try
        {
            wb = new Workbook(path, loadOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error loading workbook: {ex.Message}");
            return 0;
        }

        // Record memory usage after loading
        long afterLoad = GC.GetTotalMemory(true);

        // Dispose the workbook to release resources
        wb.Dispose();

        // Return the difference in bytes
        return afterLoad - beforeLoad;
    }
}
