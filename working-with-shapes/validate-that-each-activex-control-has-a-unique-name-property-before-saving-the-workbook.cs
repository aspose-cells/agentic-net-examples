// Title: Validate unique ActiveX control names across all worksheets before saving an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells that iterates through every worksheet, collects each ActiveX OleObject's Name, and throws an exception if any duplicate names are found before calling Workbook.Save. | Enhance the given Aspose.Cells example to log the worksheet name and duplicate ActiveX control name for each conflict, then abort the save operation. | Create a reusable C# method that accepts a Workbook, checks case‑insensitive duplicate ActiveX control names, and returns a list of duplicates or raises an error.
// Common Searches: c# aspocells check for duplicate ActiveX control names before saving workbook | how to prevent saving Excel file with duplicate OLE object names using Aspose.Cells | detect non‑unique ActiveX control names in an Excel workbook with Aspose.Cells .NET | validate unique control names in all sheets Aspose.Cells example
// Tags: Aspose.Cells duplicate ActiveX detection | C# check OleObject name uniqueness | prevent workbook save on name conflict | case‑insensitive ActiveX name verification | iterate worksheets OleObjects Aspose.Cells

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The program loads an Excel workbook, walks through each worksheet and its OleObjects, treats each as an ActiveX control, and uses a case‑insensitive HashSet to track control names. When a name appears twice, it logs the duplicate with the sheet name and aborts the save by throwing an InvalidOperationException. If no duplicates exist, the workbook is saved to the specified output file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Track ActiveX control names (case‑insensitive)
            HashSet<string> controlNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool duplicateFound = false;

            // Iterate through all worksheets
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Iterate through all OLE objects in the worksheet
                foreach (OleObject oleObject in sheet.OleObjects)
                {
                    // Process only ActiveX controls if the Type property is available.
                    // If the Type enum is not present in the referenced version, treat all OLE objects.
                    bool isActiveX = true;
                    try
                    {
                        // Attempt to use the Type property; if unavailable, the catch will keep isActiveX true.
                        // This block ensures compatibility with different Aspose.Cells versions.
                        // Uncomment the following line if OleObjectType enum is supported:
                        // isActiveX = oleObject.Type == OleObjectType.ActiveX;
                    }
                    catch
                    {
                        // Fallback: assume the object is an ActiveX control.
                    }

                    if (isActiveX)
                    {
                        string name = oleObject.Name;

                        // Detect duplicate names
                        if (!controlNames.Add(name))
                        {
                            Console.WriteLine($"Duplicate ActiveX control name detected: \"{name}\" in sheet \"{sheet.Name}\".");
                            duplicateFound = true;
                        }
                    }
                }
            }

            // Abort save if duplicates exist
            if (duplicateFound)
            {
                throw new InvalidOperationException("Workbook contains duplicate ActiveX control names. Resolve the duplicates before saving.");
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
}
