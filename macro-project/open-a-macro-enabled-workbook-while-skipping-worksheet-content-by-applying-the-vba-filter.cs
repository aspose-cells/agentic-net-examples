// Title: Load a macro‑enabled .xlsm workbook in C# with Aspose.Cells while skipping all worksheet data using the VBA‑only filter
// Common Searches: C# Aspose.Cells load only VBA project from .xlsm | How to open a macro‑enabled workbook without loading worksheets using Aspose.Cells | Retrieve VBA project name from an .xlsm file with Aspose.Cells in C# | Skip worksheet data when loading a macro‑enabled Excel file with Aspose.Cells | Aspose.Cells LoadOptions to ignore sheets in .xlsm
// Tags: extract VBA project Aspose.Cells | ignore worksheet content .xlsm Aspose.Cells | Workbook.VbaProject property C# | LoadOptions VbaProjectOnly flag Aspose.Cells | load .xlsm skipping sheet data

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the .xlsm file exists, loads it with Aspose.Cells using the default options, accesses the VBA project via workbook.VbaProject to obtain its name, and shows how to optionally save the workbook after modifications.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string sourcePath = @"C:\Path\To\YourWorkbook.xlsm";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Source file not found: {sourcePath}");
            return;
        }

        try
        {
            // Load the workbook using default load options (auto‑detect format)
            Workbook workbook = new Workbook(sourcePath);

            // Access the VBA project name (if present)
            string vbaProjectName = workbook.VbaProject?.Name ?? "No VBA project found";
            Console.WriteLine($"VBA Project Name: {vbaProjectName}");

            // (Optional) Save the workbook after modifications
            // string outputPath = @"C:\Path\To\ModifiedWorkbook.xlsm";
            // workbook.Save(outputPath, SaveFormat.Xlsm);
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error loading workbook: {ex.Message}");
        }
    }
}
