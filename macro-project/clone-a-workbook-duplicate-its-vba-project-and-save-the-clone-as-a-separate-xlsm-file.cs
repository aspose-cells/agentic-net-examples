// Title: Clone an XLSM workbook and duplicate its VBA project in C# with Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads a macro‑enabled Excel file, clones it via a MemoryStream to keep all worksheets, styles, and VBA modules, and saves the clone as a new .xlsm file using Aspose.Cells. | Show how to verify the source XLSM exists, handle exceptions, and preserve the VBA project when copying the workbook with Aspose.Cells in .NET. | Create a reusable method that accepts source and destination paths and returns a cloned workbook with its VBA project intact using Aspose.Cells SaveFormat.Xlsm.
// Common Searches: C# Aspose.Cells clone macro enabled workbook preserving VBA macros | How to copy an .xlsm file with VBA project using Aspose.Cells .NET | Duplicate Excel workbook with macros via MemoryStream Aspose.Cells example | Save cloned workbook as Xlsm format in C# Aspose.Cells | Clone Excel file and retain VBA code programmatically with Aspose
// Tags: clone workbook with VBA using Aspose.Cells | save as Xlsm format Aspose.Cells C# | memory stream workbook duplication Aspose.Cells | preserve VBA project Aspose.Cells .NET | macro-enabled Excel file copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCloneExample
{
    // The program checks for the source XLSM file, loads it with Aspose.Cells, clones the workbook through a MemoryStream to retain worksheets, styles, and VBA macros, and then saves the cloned workbook as a separate macro‑enabled .xlsm file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string sourcePath = "source.xlsm";
                const string destPath = "cloned_copy.xlsm";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the original macro-enabled workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Clone the workbook via a memory stream to preserve all content (worksheets, styles, VBA, etc.)
                Workbook clonedWorkbook;
                using (MemoryStream ms = new MemoryStream())
                {
                    sourceWorkbook.Save(ms, SaveFormat.Xlsm);
                    ms.Position = 0;
                    clonedWorkbook = new Workbook(ms);
                }

                // Save the cloned workbook as a separate macro-enabled file
                clonedWorkbook.Save(destPath, SaveFormat.Xlsm);
                Console.WriteLine($"Workbook cloned successfully to '{destPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
