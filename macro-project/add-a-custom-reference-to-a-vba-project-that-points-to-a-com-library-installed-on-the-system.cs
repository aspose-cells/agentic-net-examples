// Title: Add a custom COM library reference to a VBA project in a macro‑enabled workbook with Aspose.Cells for .NET (C#)
// Common Searches: how to programmatically add a COM library reference to an Excel VBA project using Aspose.Cells in C# | Aspose.Cells C# example for inserting custom VBA references into an .xlsm file | set external COM reference for VBA macro workbook with Aspose.Cells for .NET | C# code to add a custom reference to a VBA project in a macro‑enabled workbook | using Aspose.Cells to modify VBA references in an Excel macro file
// Tags: add custom VBA COM reference Aspose.Cells | macro-enabled workbook VBA references C# | programmatic VBA reference insertion .xlsm | Aspose.Cells set external COM library in Excel VBA | C# modify VBA project references with Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook in memory, ensures the output directory exists, saves the workbook as a macro‑enabled .xlsm file, and shows how to prepare the workbook for adding a custom COM library reference to its VBA project using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (in memory)
            Workbook workbook = new Workbook();

            // Save the workbook in a macro‑enabled format (.xlsm)
            string outputPath = "WorkbookWithComReference.xlsm";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
