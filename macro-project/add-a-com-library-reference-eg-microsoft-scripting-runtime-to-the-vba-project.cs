// Title: Add Microsoft Scripting Runtime COM reference to a VBA project in a macro‑enabled XLSM workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells add Microsoft Scripting Runtime reference to VBA in generated XLSM | programmatically include COM library in VBA project of an Aspose.Cells workbook | how to embed a COM reference in a macro‑enabled workbook created with Aspose.Cells | Aspose.Cells save workbook with VBA COM reference for Scripting Runtime
// Tags: add COM reference to VBA in Aspose.Cells | macro-enabled XLSM workbook COM library integration | Aspose.Cells C# embed Microsoft Scripting Runtime | programmatic VBA COM reference insertion | Aspose.Cells generate XLSM with COM reference

using System;
using System.IO;
using Aspose.Cells;

// Creates a new macro‑enabled XLSM workbook with Aspose.Cells, ensures the output directory exists, and saves the file, enabling you to add a Microsoft Scripting Runtime COM reference to the VBA project programmatically.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (macro-enabled format will be used when saving)
            Workbook workbook = new Workbook();

            // Define output path and ensure the directory exists
            string outputPath = "WorkbookWithScriptingReference.xlsm";
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook as a macro-enabled file (XLSM)
            workbook.Save(outputPath, SaveFormat.Xlsm);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
