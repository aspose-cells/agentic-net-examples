using System;
using System.IO;
using Aspose.Cells;

namespace MyApp
{
    class Program
    {
        static void Main()
        {
            // Path to the massive XLSX file
            string inputPath = "largeFile.xlsx";

            // Path where the processed workbook will be saved
            string outputPath = "processed.xlsx";

            try
            {
                // Prevent FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load options with low memory usage
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook using the load options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example operation: access the first worksheet and read a cell value
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

                // Save the processed workbook
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}