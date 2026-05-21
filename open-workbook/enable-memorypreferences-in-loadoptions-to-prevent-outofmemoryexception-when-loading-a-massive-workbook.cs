using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryPreferenceDemo
{
    class Program
    {
        static void Main()
        {
            // Path to the massive workbook that may cause OutOfMemoryException
            string inputPath = "MassiveWorkbook.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            try
            {
                // Create LoadOptions and enable memory preference mode
                LoadOptions loadOptions = new LoadOptions
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook using the LoadOptions with memory optimization
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example operation: read a value from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

                // Save the workbook after processing (optional)
                string outputPath = "ProcessedWorkbook.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);

                Console.WriteLine("Workbook loaded with MemoryPreference and saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}