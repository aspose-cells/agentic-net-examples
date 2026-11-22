// Title: Remove temporary custom document properties prefixed with "Temp_" from an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells delete custom document properties with a specific prefix | how to remove Excel custom properties that start with Temp_ using Aspose.Cells | clean temporary metadata from a workbook programmatically in .NET | filter and delete custom properties by name in an Excel file with Aspose.Cells | remove prefixed custom document properties from an Excel workbook in C#
// Tags: Aspose.Cells custom document properties deletion | remove properties with prefix Temp_ C# | Excel workbook metadata cleanup Aspose.Cells | Workbook.CustomDocumentProperties manipulation .NET | temporary Excel property removal programmatically

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads an Excel workbook with Aspose.Cells, iterates its CustomDocumentProperties collection to find entries whose names begin with "Temp_", removes those properties, ensures the output directory exists, and saves the cleaned workbook to a new file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access custom document properties collection
            var customProps = workbook.CustomDocumentProperties;

            // Collect names of properties that start with "Temp_"
            List<string> namesToRemove = new List<string>();
            foreach (var prop in customProps)
            {
                // prop is of type CustomDocumentProperty
                if (prop.Name != null && prop.Name.StartsWith("Temp_", StringComparison.OrdinalIgnoreCase))
                {
                    namesToRemove.Add(prop.Name);
                }
            }

            // Remove identified temporary properties
            foreach (string name in namesToRemove)
            {
                customProps.Remove(name);
            }

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
