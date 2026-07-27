// Title: Load a Massive XLSX with Low Memory Using Aspose.Cells MemoryPreference (C#)
// Description: Shows how to set LoadOptions.MemorySetting to MemoryPreference (low‑memory mode) before opening a large XLSX file with Aspose.Cells for .NET, minimizing RAM consumption while accessing worksheets and saving the workbook.
// Keywords: Aspose.Cells | MemoryPreference | LoadOptions | low memory mode | large XLSX | C# | .NET | reduce RAM usage | open massive workbook | Excel memory optimization | GitHub example
// Common Searches: Aspose.Cells set MemoryPreference before loading large Excel | C# load big XLSX with low memory usage | LoadOptions MemorySetting low‑memory example | reduce RAM when opening massive workbook Aspose.Cells | how to use MemoryPreference in Aspose.Cells .NET
// Developer Intent: Configure Aspose.Cells to open a huge XLSX file in low‑memory mode using MemoryPreference.
// Use Cases: Process multi‑gigabyte Excel files on machines with limited RAM. | Read worksheet metadata (e.g., first sheet name) without loading the entire workbook into memory. | Integrate low‑memory loading into batch jobs that handle dozens of large spreadsheets.
// AI Prompts: Generate C# code that opens a large XLSX with Aspose.Cells using MemoryPreference, iterates all worksheets, and prints each sheet name. | Provide an example that modifies a cell in a massive workbook loaded in low‑memory mode and saves the file without high RAM consumption. | Explain how to combine MemoryPreference with other LoadOptions such as LoadFormat and Password for optimal performance on big Excel datasets.

using System;
using Aspose.Cells;

namespace AsposeCellsMemoryPreferenceDemo
{
    // Shows how to set LoadOptions.MemorySetting to MemoryPreference (low‑memory mode) before opening a large XLSX file with Aspose.Cells for .NET, minimizing RAM consumption while accessing worksheets and saving the workbook.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the massive XLSX file
            string inputFile = "largeFile.xlsx";

            // Path where the processed workbook will be saved
            string outputFile = "processed.xlsx";

            // Create LoadOptions and set the memory mode to MemoryPreference
            // This reduces RAM usage when loading large workbooks
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.MemorySetting = MemorySetting.MemoryPreference;

            // Load the workbook using the configured LoadOptions
            Workbook workbook = new Workbook(inputFile, loadOptions);

            // Example operation: display the name of the first worksheet
            Console.WriteLine("First worksheet name: " + workbook.Worksheets[0].Name);

            // Save the workbook (you can choose any format you need)
            workbook.Save(outputFile, SaveFormat.Xlsx);

            Console.WriteLine("Workbook processed and saved with MemoryPreference setting.");
        }
    }
}
