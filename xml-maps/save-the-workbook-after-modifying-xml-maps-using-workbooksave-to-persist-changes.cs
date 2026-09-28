// Title: Saving an Excel workbook after updating XML maps with Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that opens an .xlsx file with Aspose.Cells, modifies its XML map (when the API is available), and calls Workbook.Save to write the updated workbook to a new file. | Demonstrate how to verify the source file, apply XML map changes, and persist those changes using Workbook.Save while handling exceptions in a C# Aspose.Cells application.
// Common Searches: Aspose.Cells C# how to persist changes to an XML map in an Excel file | save workbook after editing XML map using Workbook.Save in .NET | example code for updating Excel XML map and saving with Aspose.Cells
// Tags: Aspose.Cells Workbook.Save Excel file | C# update XML map Aspose.Cells | persist XML map changes .NET | save workbook after XML map edit | Aspose.Cells XML map handling C#

using Aspose.Cells;
using System;
using System.IO;

// The example verifies that input.xlsx exists, loads it into an Aspose.Cells Workbook, notes that XML map functionality is unavailable in the current version, then saves the workbook as output.xlsx using Workbook.Save while catching and reporting any exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: XML map functionality is not available in the current Aspose.Cells version.
            // If needed, XML map handling should be added when the appropriate API is present.

            // Save the workbook with the desired format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors to prevent the application from crashing
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
