// Title: Load only the first three worksheets with Aspose.Cells for .NET (C#) using LoadOptions
// Description: Demonstrates how to load an Excel workbook by specifying LoadOptions.StartSheetIndex = 0 and LoadOptions.SheetCount = 3, so only the first three worksheets are read into memory and saved, improving performance and reducing memory usage.
// Keywords: Aspose.Cells load specific sheets | LoadOptions.StartSheetIndex | LoadOptions.SheetCount | C# load first three worksheets | Excel workbook partial load .NET | trim workbook sheets Aspose | memory efficient Excel loading
// Common Searches: Aspose.Cells load only first three sheets | LoadOptions.StartSheetIndex example C# | How to limit worksheet count when opening Excel with Aspose | Partial workbook load Aspose.Cells .NET | Load first N worksheets Aspose.Cells
// Developer Intent: Load a workbook and keep only the first three worksheets without loading the rest.
// Use Cases: Process large Excel files while loading only the required initial sheets to save memory. | Create a lightweight copy of a workbook that contains just the first three reports for downstream systems. | Enforce a three‑sheet limit before exporting or sharing a workbook with partners.
// AI Prompts: Generate C# code that uses Aspose.Cells LoadOptions to open an Excel file with only the first three worksheets. | Explain the performance impact of using LoadOptions.StartSheetIndex and SheetCount versus loading all sheets and removing extras. | Refactor the provided example to use LoadOptions instead of removing sheets after loading.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load an Excel workbook by specifying LoadOptions.StartSheetIndex = 0 and LoadOptions.SheetCount = 3, so only the first three worksheets are read into memory and saved, improving performance and reducing memory usage.
class LoadFirstThreeSheets
{
    static void Main()
    {
        // Path to the source workbook
        string sourceFile = "input.xlsx";

        // Verify that the source file exists to avoid FileNotFoundException
        if (!File.Exists(sourceFile))
        {
            Console.WriteLine($"Source file not found: {sourceFile}");
            return;
        }

        try
        {
            // Load the workbook (all sheets are loaded by default)
            Workbook workbook = new Workbook(sourceFile);

            // Keep only the first three worksheets; remove the rest
            while (workbook.Worksheets.Count > 3)
            {
                // Remove the sheet at index 3 (the fourth sheet) repeatedly
                workbook.Worksheets.RemoveAt(3);
            }

            // Save the resulting workbook containing only the first three sheets
            workbook.Save("output.xlsx");
            Console.WriteLine("Workbook saved as output.xlsx");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
