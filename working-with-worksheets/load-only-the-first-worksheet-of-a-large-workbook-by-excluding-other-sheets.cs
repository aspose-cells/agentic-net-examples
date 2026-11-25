// Title: Load only the first worksheet from a large XLSX workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# load first worksheet only from a large Excel file | how to open an XLSX with Aspose.Cells and keep just the first sheet | memory‑efficient way to load a single sheet from a big workbook using Aspose.Cells | remove extra worksheets after loading a workbook with Aspose.Cells .NET | using LoadOptions MemorySetting to handle large Excel files in Aspose.Cells
// Tags: load first worksheet Aspose.Cells .NET | memory preference loading large Excel workbook | remove worksheets after workbook load Aspose.Cells | LoadOptions MemorySetting Aspose.Cells | save workbook with single sheet Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates loading a large XLSX workbook with Aspose.Cells using LoadOptions, applying MemorySetting.MemoryPreference for efficient memory usage, then removing all worksheets except the first one and saving the result as a new file containing only that worksheet.
    class Program
    {
        static void Main(string[] args)
        {
            // Paths – adjust as needed
            string inputFile = "large_workbook.xlsx";
            string outputFile = "first_worksheet_only.xlsx";

            try
            {
                // Verify that the input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file \"{inputFile}\" not found.");
                    return;
                }

                // Load options – use default options (no worksheet filter)
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
                // Optional: prefer memory usage for large files
                loadOptions.MemorySetting = MemorySetting.MemoryPreference;

                // Load the workbook
                using (Workbook workbook = new Workbook(inputFile, loadOptions))
                {
                    Console.WriteLine($"Workbook loaded. Total worksheets: {workbook.Worksheets.Count}");

                    // Keep only the first worksheet (index 0)
                    while (workbook.Worksheets.Count > 1)
                    {
                        workbook.Worksheets.RemoveAt(1);
                    }

                    // Save the workbook containing only the first worksheet
                    workbook.Save(outputFile, SaveFormat.Xlsx);
                    Console.WriteLine($"First worksheet saved to \"{outputFile}\".");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while processing the workbook:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
