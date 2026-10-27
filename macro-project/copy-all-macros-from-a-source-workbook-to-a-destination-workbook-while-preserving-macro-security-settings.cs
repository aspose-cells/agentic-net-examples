// Title: Copy VBA macros from a source .xlsm workbook to another .xlsm workbook with Aspose.Cells for .NET (C#)
// Common Searches: how to copy VBA macros between two .xlsm files using Aspose.Cells C# | Aspose.Cells C# example for duplicating a macro‑enabled workbook | preserve macro security settings when saving an .xlsm workbook with Aspose.Cells | copy VBA project from one workbook to another in a .NET application | clone a macro‑enabled Excel file programmatically with Aspose.Cells
// Tags: copy VBA macros Aspose.Cells C# | save macro‑enabled workbook Aspose.Cells | preserve macro security Aspose.Cells | duplicate VBA project between workbooks | clone .xlsm file using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Vba; // Required for VBA project APIs

// The example checks that the source .xlsm file exists, loads it into an Aspose.Cells Workbook, and saves it to a new .xlsm file using SaveFormat.Xlsm. This operation copies all VBA macros and retains the original macro security settings, with error handling via a try‑catch block.
class MacroCopyExample
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsm";
            const string destPath = "DestinationWorkbook.xlsm";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook that contains macros
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Save the workbook to the destination path in macro‑enabled format.
            // This copies all content, including VBA macros, from the source workbook.
            sourceWorkbook.Save(destPath, SaveFormat.Xlsm);

            Console.WriteLine($"Macros copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
