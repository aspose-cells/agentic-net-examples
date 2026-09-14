// Title: Remove all custom document properties prefixed with "Temp_" from an Excel workbook using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, scans the CustomDocumentProperties collection, deletes every property whose name begins with "Temp_", and saves the workbook. | Demonstrate how to iterate backwards through Aspose.Cells CustomDocumentProperties in C# to safely remove matching entries without collection errors. | Write a C# snippet that logs each removed temporary custom property name before saving the updated workbook with Aspose.Cells.
// Common Searches: Aspose.Cells C# remove custom document properties that start with a specific prefix | How to delete temporary metadata from an Excel file using Aspose.Cells .NET | Iterating backwards through CustomDocumentProperties collection in Aspose.Cells | Programmatically clean up custom properties in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells remove custom document properties | C# delete custom properties by prefix | Excel workbook metadata cleanup Aspose.Cells | CustomDocumentProperties backward iteration | temporary property removal .NET Excel

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook (or creates a new one), accesses its CustomDocumentProperties collection, iterates in reverse order, removes any property whose name starts with "Temp_", optionally logs removed names, and saves the cleaned workbook.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Load workbook; create a new one if the input file does not exist.
                Workbook workbook;
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access custom document properties.
                var customProps = workbook.CustomDocumentProperties;

                // Remove properties whose name starts with "Temp_".
                for (int i = customProps.Count - 1; i >= 0; i--)
                {
                    var prop = customProps[i];
                    if (!string.IsNullOrEmpty(prop.Name) && prop.Name.StartsWith("Temp_", StringComparison.Ordinal))
                    {
                        customProps.Remove(prop.Name);
                    }
                }

                // Save the modified workbook.
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
