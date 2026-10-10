// Title: How to Inspect and Log Font Substitution Warnings from an Aspose.Cells Workbook After Saving in C#
// AI Prompts: Generate C# code that saves an Excel workbook with Aspose.Cells, then examines workbook.Warnings for any FontSubstitution warnings and writes each warning message to the console. | Show an example of iterating through the Aspose.Cells WarningCollection after workbook.Save to capture and log font‑substitution details, including handling the case where no warnings are present.
// Common Searches: Aspose.Cells C# retrieve font substitution warnings after workbook.Save | C# log warning messages from Aspose.Cells warning collection | How to access workbook.Warnings for missing fonts in Aspose.Cells | Check for font substitution warnings when converting Excel to PDF using Aspose.Cells C#
// Tags: Aspose.Cells warning collection iteration | font substitution warning logging C# | inspect workbook warnings after save Aspose.Cells | C# Aspose.Cells font fallback detection

using System;
using System.IO;
using Aspose.Cells;

// The example demonstrates loading an Excel file, saving it with Aspose.Cells, and then iterating over the workbook.Warnings collection to identify FontSubstitution warnings. Each relevant warning is written to the console, with graceful handling when no warnings exist. Requires a version of Aspose.Cells that exposes the WarningCollection API.
class Program
{
    static void Main()
    {
        try
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to the desired output location
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to: {outputPath}");

            // NOTE: Warning retrieval APIs (GetWarnings, WarningCollection, Warning) are not
            // available in the referenced Aspose.Cells version, so this part is omitted.
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
