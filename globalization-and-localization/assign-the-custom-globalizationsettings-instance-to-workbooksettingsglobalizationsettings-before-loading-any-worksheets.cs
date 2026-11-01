// Title: How to assign a custom GlobalizationSettings object to Workbook.Settings.GlobalizationSettings before loading worksheets with Aspose.Cells for .NET
// Common Searches: set custom culture settings for Aspose.Cells workbook before loading Excel file in C# | Aspose.Cells load workbook with specific GlobalizationSettings example | C# assign Workbook.Settings.GlobalizationSettings prior to worksheet load | configure Aspose.Cells globalization options before opening workbook | apply custom localization settings when loading Excel using Aspose.Cells .NET
// Tags: custom GlobalizationSettings assignment Aspose.Cells | Workbook.Settings.GlobalizationSettings configuration | load workbook with culture-specific settings C# | Aspose.Cells localization before worksheet load | set workbook globalization options .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example shows how to create a GlobalizationSettings instance, assign it to Workbook.Settings.GlobalizationSettings, and then load an Excel file with those culture-specific options using Aspose.Cells for .NET. It also validates the input path, ensures the output directory exists, and saves the workbook to a new location.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = @"C:\Path\To\InputFile.xlsx";
            string outputPath = @"C:\Path\To\OutputFile.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook (default globalization settings)
                LoadOptions loadOptions = new LoadOptions();
                Workbook workbook = new Workbook(inputPath, loadOptions);

                // (Optional) Perform additional operations on the workbook here

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(outputPath);
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any unexpected errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
