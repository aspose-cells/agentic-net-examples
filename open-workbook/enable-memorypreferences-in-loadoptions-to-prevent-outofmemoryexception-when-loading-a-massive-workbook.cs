// Title: Loading a massive XLSX workbook in C# with Aspose.Cells using MemoryPreference to avoid OutOfMemoryException
// Common Searches: C# Aspose.Cells load large XLSX file with MemoryPreference to reduce memory usage | How to set LoadOptions.MemorySetting for huge Excel workbooks in Aspose.Cells | Prevent OutOfMemoryException when opening big Excel files with Aspose.Cells .NET | Example of using LoadOptions MemoryPreference in Aspose.Cells C# code | Best practice for low‑memory loading of massive workbooks in Aspose.Cells
// Tags: Aspose.Cells LoadOptions MemoryPreference | load large XLSX low memory Aspose.Cells | prevent OutOfMemoryException Aspose.Cells | C# memory optimization Aspose.Cells | massive workbook loading Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a LoadOptions object for XLSX files, sets its MemorySetting to MemoryPreference to lower memory consumption, loads the specified workbook with these options, and saves it to a new file; it also generates a placeholder workbook if the input file is missing.
class Program
{
    static void Main()
    {
        const string inputPath = "MassiveWorkbook.xlsx";
        const string outputPath = "MassiveWorkbook_Output.xlsx";

        try
        {
            // Ensure the input file exists; create a placeholder if it does not.
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Creating an empty workbook as a placeholder.");
                var placeholder = new Workbook();
                placeholder.Save(inputPath);
            }

            // Configure load options to use memory preferences, reducing OOM risk.
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                MemorySetting = MemorySetting.MemoryPreference
            };

            // Load the workbook with the specified memory settings.
            var workbook = new Workbook(inputPath, loadOptions);

            // (Optional) Perform any required operations on the workbook here.

            // Save the workbook after processing.
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
