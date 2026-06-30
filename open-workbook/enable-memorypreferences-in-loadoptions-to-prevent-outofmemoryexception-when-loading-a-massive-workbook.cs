using System;
using Aspose.Cells;

// Author: Example demonstrating MemoryPreference to handle large workbooks
class Program
{
    static void Main()
    {
        // Create LoadOptions and set memory mode to prefer lower memory usage
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
        loadOptions.MemorySetting = MemorySetting.MemoryPreference;

        // Load the massive workbook with the configured options
        Workbook workbook = new Workbook("LargeWorkbook.xlsx", loadOptions);

        // Sample operation: output the name of the first worksheet
        Worksheet firstSheet = workbook.Worksheets[0];
        Console.WriteLine($"First worksheet name: {firstSheet.Name}");

        // Save the workbook (optional, can be to a different format or location)
        workbook.Save("LargeWorkbook_Processed.xlsx", SaveFormat.Xlsx);
    }
}