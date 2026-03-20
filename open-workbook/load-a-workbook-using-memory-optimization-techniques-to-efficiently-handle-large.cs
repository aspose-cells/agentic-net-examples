using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizedLoad
{
    class Program
    {
        static void Main()
        {
            // Path to the large Excel file to be loaded
            string inputPath = "LargeDataSet.xlsx";

            // Path where the processed workbook will be saved
            string outputPath = "LargeDataSet_Optimized.xlsx";

            // Create LoadOptions and configure memory‑optimized settings
            LoadOptions loadOptions = new LoadOptions();
            // Prefer memory‑efficient mode (compact data representation)
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;
            // When we only need to read data and not save back the original unparsed parts,
            // disabling this can improve performance
            loadOptions.KeepUnparsedData = false;

            // Open the source file as a stream (avoids loading the whole file into memory at once)
            using (FileStream fileStream = new FileStream(inputPath, FileMode.Open, FileAccess.Read))
            {
                // Load the workbook using the stream constructor that accepts LoadOptions
                Workbook workbook = new Workbook(fileStream, loadOptions);

                // Example operation: read the first worksheet name and first cell value
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine($"Worksheet: {sheet.Name}");
                Console.WriteLine($"A1 Value: {sheet.Cells["A1"].StringValue}");

                // Save the workbook (the same memory‑optimized settings are retained)
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to '{outputPath}' with memory‑optimized settings.");
            }
        }
    }
}