// Title: Load Large Excel Files with Optimized Memory Using Aspose.Cells for .NET
// Description: Demonstrates how to activate Aspose.Cells LoadOptions.MemorySetting.MemoryPreference to lower RAM consumption when opening workbooks that contain thousands of rows, read a cell value, and optionally save the file.
// Keywords: Aspose.Cells | LoadOptions | MemorySetting.MemoryPreference | optimized memory | large workbook | .NET | C# | reduce memory usage | Excel performance
// Common Searches: Aspose.Cells enable optimized memory | load large Excel workbook low memory .NET | MemorySetting.MemoryPreference example C# | Aspose.Cells reduce RAM usage when loading | how to open big Excel file with Aspose.Cells
// Developer Intent: Configure LoadOptions.MemorySetting to MemoryPreference to load a massive Excel workbook with a smaller memory footprint in C#.
// Use Cases: Open a multi‑megabyte spreadsheet containing tens of thousands of rows on a memory‑constrained server. | Read or validate specific cells after loading without fully materializing the workbook in memory. | Process and re‑save large spreadsheets while keeping RAM usage minimal.
// AI Prompts: Provide a C# snippet that uses Aspose.Cells LoadOptions with MemorySetting.MemoryPreference to open a large .xlsx file. | Explain the trade‑offs of MemorySetting.MemoryPreference versus default loading in Aspose.Cells and recommend best practices for performance and memory management.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to activate Aspose.Cells LoadOptions.MemorySetting.MemoryPreference to lower RAM consumption when opening workbooks that contain thousands of rows, read a cell value, and optionally save the file.
class Program
{
    static void Main()
    {
        // Path to the workbook that contains thousands of rows
        string inputPath = "large_dataset.xlsx";

        try
        {
            // Ensure the input file exists; create a minimal workbook if it does not
            if (!File.Exists(inputPath))
            {
                Workbook tempWb = new Workbook();
                tempWb.Worksheets[0].Cells["A1"].PutValue("Sample");
                tempWb.Save(inputPath, SaveFormat.Xlsx);
            }

            // Configure LoadOptions to use optimized memory mode
            LoadOptions loadOptions = new LoadOptions
            {
                // MemoryPreference reduces memory usage at the cost of some performance overhead
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook using the configured LoadOptions
            Workbook workbook = new Workbook(inputPath, loadOptions);

            // Example operation: read a cell value to verify the workbook is loaded
            Worksheet sheet = workbook.Worksheets[0];
            Console.WriteLine("Cell A1 value: " + sheet.Cells["A1"].StringValue);

            // Save the workbook (optional, demonstrates the full lifecycle)
            workbook.Save("optimized_output.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
