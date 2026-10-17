// Title: Delete custom document properties prefixed with "Temp_" from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells remove custom document properties that start with Temp_ | How to clean temporary metadata from an Excel file using Aspose.Cells | Aspose.Cells delete custom properties by name prefix in .NET | Programmatically remove Temp_ properties from a workbook with Aspose.Cells
// Tags: Aspose.Cells custom document property deletion | C# filter and delete Excel custom properties | temporary workbook metadata cleanup Aspose | dynamic enumeration of custom properties .NET | save workbook after property removal

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook, iterates its custom document properties, identifies those whose names begin with "Temp_", removes the temporary entries, and saves the cleaned file, with checks for missing input and exception handling.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access custom document properties using dynamic to avoid compile‑time type issues
            dynamic customProps = workbook.CustomDocumentProperties;

            // Collect names of temporary properties (those starting with "Temp_")
            List<string> namesToRemove = new List<string>();
            foreach (dynamic prop in customProps)
            {
                string propName = prop.Name as string;
                if (!string.IsNullOrEmpty(propName) && propName.StartsWith("Temp_"))
                {
                    namesToRemove.Add(propName);
                }
            }

            // Remove the identified temporary properties
            foreach (string name in namesToRemove)
            {
                customProps.Remove(name);
            }

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
