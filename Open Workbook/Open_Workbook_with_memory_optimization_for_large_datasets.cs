using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryOptimizedDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the large Excel file to be loaded
            string inputPath = "LargeDataFile.xlsx";

            // If the file does not exist, create a simple workbook for demonstration purposes
            if (!File.Exists(inputPath))
            {
                var tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Demo");
                tempWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Create LoadOptions and set memory mode to MemoryPreference
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example operation: read the value of cell A1 from the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook after processing (optional)
            string outputPath = "ProcessedLargeDataFile.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);

            Console.WriteLine("Workbook processed and saved with memory‑optimized settings.");
        }
    }
}