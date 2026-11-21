// Title: Clone an XLSM workbook with its VBA macros and save as a new file using Aspose.Cells for .NET
// Common Searches: how to duplicate an XLSM workbook with macros using Aspose.Cells C# | Aspose.Cells clone workbook preserving VBA project example | save cloned workbook to memory stream then to file Aspose.Cells | copy Excel macro-enabled workbook programmatically in .NET | C# code to clone macro-enabled workbook with Aspose.Cells
// Tags: clone workbook with VBA macros Aspose.Cells | save workbook to memory stream C# | duplicate macro-enabled Excel file using Aspose.Cells | load and save XLSM with VBA project Aspose.Cells | programmatic Excel workbook cloning .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled XLSM workbook, clones it by saving to a MemoryStream and reloading, and then saves the cloned workbook—including its VBA project—as a separate XLSM file, with error handling for missing files.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsm";
        const string outputPath = "cloned_workbook.xlsm";

        // Verify source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the original XLSM workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Clone the workbook by saving to a memory stream and reloading
            Workbook clonedWorkbook;
            using (MemoryStream ms = new MemoryStream())
            {
                sourceWorkbook.Save(ms, SaveFormat.Xlsm);
                ms.Position = 0;
                clonedWorkbook = new Workbook(ms);
            }

            // Save the cloned workbook as a separate XLSM file
            clonedWorkbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Cloned workbook saved to: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
