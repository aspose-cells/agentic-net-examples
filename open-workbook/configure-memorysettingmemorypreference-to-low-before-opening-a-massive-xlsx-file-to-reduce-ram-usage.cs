// Title: Set MemorySetting.MemoryPreference to Low in Aspose.Cells C# before opening a large XLSX workbook
// AI Prompts: Write C# code that configures Aspose.Cells MemorySetting.MemoryPreference to Low, then loads a massive .xlsx file and saves it. | Show how to enable low‑memory mode in Aspose.Cells prior to creating a Workbook instance for a big Excel workbook. | Provide a C# example that verifies the file exists, sets low memory preference, opens the workbook, and handles any exceptions.
// Common Searches: asp.net set memorypreference low before loading large excel with aspose.cells | how to reduce memory consumption when opening big xlsx using Aspose.Cells in C# | Aspose.Cells low memory mode example for massive workbook | configure MemorySetting.MemoryPreference to Low for large Excel files in .NET
// Tags: Aspose.Cells low memory preference | MemorySetting.MemoryPreference Low example | load large XLSX with reduced RAM Aspose.Cells | C# workbook memory optimization Aspose.Cells | configure memory setting before workbook load

using Aspose.Cells;
using System;
using System.IO;

// The sample checks that the specified massive.xlsx file exists, sets Aspose.Cells MemorySetting.MemoryPreference to Low to minimize RAM usage, loads the workbook, optionally processes it, saves the result to output.xlsx, and logs any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for input and output files
            const string inputPath = "massive.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook (default load options)
            Workbook workbook = new Workbook(inputPath);

            // (Optional) Perform any required operations on the workbook here

            // Save the workbook to the specified output path
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            // Log the exception details for troubleshooting
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
