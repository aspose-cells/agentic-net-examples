// Title: Check for duplicate ActiveX (OLE) control names in an Aspose.Cells workbook before saving (C#)
// AI Prompts: Write a C# method with Aspose.Cells that iterates all worksheets and throws an exception when two OleObject controls share the same Name. | Add a pre‑save validation step to an Aspose.Cells workbook that verifies each ActiveX control has a unique Name across the entire file. | Generate code that loads an Excel file, validates unique OleObject names case‑insensitively, and saves only if no duplicates are found.
// Common Searches: Aspose.Cells C# how to ensure ActiveX control names are unique before saving workbook | detect duplicate OleObject names in Excel using Aspose.Cells library | C# validate unique ActiveX control Name property across all worksheets with Aspose.Cells | throw exception for duplicate OLE object names in Aspose.Cells workbook
// Tags: validate unique OleObject names Aspose.Cells | duplicate ActiveX control detection C# | pre‑save workbook validation Aspose.Cells | case‑insensitive control name check Excel | throw exception on duplicate OLE objects

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// The example loads an Excel workbook with Aspose.Cells, scans every worksheet for OleObject (ActiveX) controls, tracks their Name properties in a case‑insensitive HashSet, and throws an InvalidOperationException if a duplicate name is found, ensuring the workbook is saved only after name uniqueness is confirmed.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Validate that each ActiveX (OLE) control has a unique Name property
            ValidateActiveXControlNames(workbook);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    static void ValidateActiveXControlNames(Workbook workbook)
    {
        // Track encountered control names (case‑insensitive)
        HashSet<string> controlNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Iterate through all OLE objects (ActiveX controls) on the worksheet
            foreach (OleObject ole in sheet.OleObjects)
            {
                string name = ole.Name;

                // Skip objects without a name
                if (string.IsNullOrEmpty(name))
                    continue;

                // Detect duplicate names
                if (!controlNames.Add(name))
                {
                    throw new InvalidOperationException(
                        $"Duplicate ActiveX control name detected: '{name}' in worksheet '{sheet.Name}'.");
                }
            }
        }
    }
}
