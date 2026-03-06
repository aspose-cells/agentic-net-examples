using System;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizationDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Create LoadOptions and set the memory usage mode to MemoryPreference
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Open the workbook with the specified load options
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Example operation: display the number of worksheets
            Console.WriteLine("Number of worksheets: " + workbook.Worksheets.Count);

            // Save the workbook (the memory setting is retained during processing)
            workbook.Save("output_optimized.xlsx", SaveFormat.Xlsx);

            // Release resources
            workbook.Dispose();
        }
    }
}