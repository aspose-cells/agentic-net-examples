using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Author: Aspose.Cells .NET example – demonstrates enabling optimized memory usage when loading large workbooks
    class LoadWorkbookWithOptimizedMemory
    {
        static void Main()
        {
            // Create LoadOptions and set memory mode to prefer optimized memory usage
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference; // enables memory‑optimized loading

            // Load the workbook using the configured LoadOptions
            // Replace "LargeDataWorkbook.xlsx" with the path to your source file
            Workbook workbook = new Workbook("LargeDataWorkbook.xlsx", loadOptions);

            // (Optional) Perform any workbook operations here
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine($"First cell value: {sheet.Cells["A1"].StringValue}");

            // Save the workbook after processing
            workbook.Save("OptimizedMemoryOutput.xlsx", SaveFormat.Xlsx);
        }
    }
}