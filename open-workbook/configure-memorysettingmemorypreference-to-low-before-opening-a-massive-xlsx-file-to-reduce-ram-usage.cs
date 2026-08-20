// Title: Load a Large XLSX Workbook with Low Memory Using Aspose.Cells MemoryPreference (C#)
// Description: Shows how to set Aspose.Cells LoadOptions.MemorySetting to MemoryPreference for opening massive XLSX files with minimal RAM, including file‑existence validation, placeholder workbook creation, cell reading, and saving the result.
// Keywords: Aspose.Cells | MemoryPreference | low memory loading | large XLSX | LoadOptions | C# example | reduce RAM | placeholder workbook | massive file | memory setting
// Common Searches: Aspose.Cells low memory mode C# | How to open large XLSX with reduced RAM using Aspose.Cells | Set MemorySetting.MemoryPreference before loading workbook | Create placeholder workbook when file missing Aspose.Cells | LoadOptions memory setting example
// Developer Intent: Configure Aspose.Cells to use the MemoryPreference setting before loading a huge XLSX workbook so that RAM consumption stays low.
// Use Cases: Process multi‑gigabyte XLSX files on machines with limited memory. | Run batch conversions in cloud functions where memory is constrained. | Validate or extract data from large spreadsheets without exhausting resources. | Generate a fallback workbook when the source file is unavailable.
// AI Prompts: Write C# code that sets LoadOptions.MemorySetting to MemoryPreference and opens a large XLSX file with Aspose.Cells. | Show how to check for a missing Excel file, create a simple placeholder workbook, then load the target workbook using low‑memory settings. | Explain the performance impact of MemoryPreference versus other memory settings in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsMemoryPreferenceDemo
{
    // Shows how to set Aspose.Cells LoadOptions.MemorySetting to MemoryPreference for opening massive XLSX files with minimal RAM, including file‑existence validation, placeholder workbook creation, cell reading, and saving the result.
    class Program
    {
        static void Main()
        {
            // Path to the massive XLSX file
            string inputPath = "massive_file.xlsx";

            try
            {
                // Verify the input file exists
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    // Create a simple placeholder workbook to allow the demo to continue
                    Workbook placeholder = new Workbook();
                    placeholder.Worksheets[0].Cells["A1"].PutValue("Placeholder");
                    string placeholderPath = "placeholder.xlsx";
                    placeholder.Save(placeholderPath, SaveFormat.Xlsx);
                    Console.WriteLine($"Created placeholder file: {placeholderPath}");
                    inputPath = placeholderPath;
                }

                // Configure load options for low RAM usage
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };

                // Load the workbook with the specified options
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // Example operation: read the value of cell A1 from the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                Console.WriteLine("A1 value: " + sheet.Cells["A1"].StringValue);

                // Save the workbook after processing (optional)
                string outputPath = "processed_file.xlsx";
                workbook.Save(outputPath, SaveFormat.Xlsx);
                Console.WriteLine("Workbook loaded with MemoryPreference and saved to: " + outputPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
