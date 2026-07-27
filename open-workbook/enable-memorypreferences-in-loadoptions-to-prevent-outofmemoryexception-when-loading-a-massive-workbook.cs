// Title: Load huge Excel workbook with MemoryPreference to avoid OutOfMemoryException – Aspose.Cells .NET
// Description: Shows how to set LoadOptions.MemorySetting = MemoryPreference in Aspose.Cells to open massive .xlsx files without exhausting memory, with a fallback workbook creation, a sample write to A1, and saving.
// Keywords: Aspose.Cells | MemoryPreference | LoadOptions | large workbook | OutOfMemoryException | .NET | Excel streaming | memory management | massive Excel file | Load large XLSX
// Common Searches: Aspose.Cells load large workbook memory preference | MemorySetting MemoryPreference C# example | prevent OutOfMemoryException Aspose.Cells | open massive Excel file with Aspose.Cells .NET | LoadOptions MemoryPreference usage guide
// Developer Intent: Open a very large Excel file in .NET with Aspose.Cells while keeping memory consumption low.
// Use Cases: Read a multi‑gigabyte .xlsx for data extraction without hitting memory limits. | Add a processing timestamp to the first sheet of a huge workbook and save it using the same memory setting. | Create a new workbook when the source file is missing, still applying MemoryPreference for consistent resource handling.
// AI Prompts: Generate C# code that reads a 5 GB Excel file with Aspose.Cells using MemoryPreference and iterates over rows safely. | Explain how to combine MemoryPreference with Aspose.Cells streaming API to process large worksheets efficiently. | Show how to switch between MemoryPreference and MemorySafe in LoadOptions based on file size thresholds.

using System;
using System.IO;
using Aspose.Cells;

// Shows how to set LoadOptions.MemorySetting = MemoryPreference in Aspose.Cells to open massive .xlsx files without exhausting memory, with a fallback workbook creation, a sample write to A1, and saving.
class Program
{
    static void Main()
    {
        // Path to the large workbook that may cause OutOfMemoryException
        string inputFile = "MassiveWorkbook.xlsx";

        // Optional: path to save the workbook after processing
        string outputFile = "MassiveWorkbook_Processed.xlsx";

        try
        {
            Workbook workbook;

            if (File.Exists(inputFile))
            {
                // Load the workbook with memory‑preference mode
                LoadOptions loadOptions = new LoadOptions
                {
                    MemorySetting = MemorySetting.MemoryPreference
                };
                workbook = new Workbook(inputFile, loadOptions);
            }
            else
            {
                // Create a new workbook if the input file is missing
                workbook = new Workbook();
                workbook.Worksheets[0].Name = "Sheet1";
            }

            // Example operation: write a timestamp to cell A1 of the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue($"Loaded on {DateTime.Now}");

            // Save the workbook (optional, demonstrates that saving works with the chosen memory setting)
            workbook.Save(outputFile, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{outputFile}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
