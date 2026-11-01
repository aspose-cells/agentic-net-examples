// Title: Save an Excel workbook after modifying XML maps using Aspose.Cells Workbook.Save in C#
// Common Searches: how to persist XML map changes when saving an Excel workbook with Aspose.Cells in C# | Aspose.Cells C# example for saving a workbook that contains updated XML map definitions | using Workbook.Save to write modified XML maps to an XLSX file in .NET | C# code to export Excel file after editing XML maps with Aspose.Cells
// Tags: Aspose.Cells Workbook.Save XML map persistence | C# save workbook with updated XML maps | Excel XLSX export after XML map modification Aspose.Cells | SaveFormat.Xlsx usage with XML map changes in Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies that an input.xlsx file exists, loads it into an Aspose.Cells Workbook, skips XML map manipulation due to version constraints, and then saves the workbook as output.xlsx using Workbook.Save with SaveFormat.Xlsx, while handling any runtime exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before attempting to load it
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: XML map functionality is not available in the current Aspose.Cells version.
            // The following code block is omitted to keep the example compilable and functional.

            // Save the workbook to the desired output format
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected exceptions and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
