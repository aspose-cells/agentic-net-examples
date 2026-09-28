// Title: Use Aspose.Cells for .NET to load an Excel workbook, remove all defined names that contain “Total”, and export the result to PDF
// AI Prompts: Write C# code with Aspose.Cells that opens an .xlsx file, finds every defined name containing the word “Total” (case‑insensitive), deletes those names, and saves the workbook as a PDF. | Generate a .NET example that filters named ranges by a keyword, removes the matching ranges, and converts the cleaned workbook to a PDF document using Aspose.Cells.
// Common Searches: asp.net remove named ranges containing 'Total' from Excel file using Aspose.Cells | convert Excel to PDF after deleting specific defined names with Aspose.Cells C# | filter workbook defined names by keyword and save as PDF in C# Aspose.Cells example
// Tags: defined name cleanup Aspose.Cells | keyword-based named range removal Aspose.Cells | Excel to PDF conversion after name cleanup Aspose.Cells | case-insensitive defined name filter Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;
using System.Linq;

// // Loads input.xlsx, deletes all defined names that include "Total" (case‑insensitive), and saves the modified workbook as output.pdf using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.pdf";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook (standard loading; LightCells API not required for this task)
            Workbook workbook = new Workbook(inputPath);

            // Find all defined names that contain the word "Total" (case‑insensitive)
            var namesToRemove = workbook.Worksheets.Names
                .Where(n => n.Text.IndexOf("Total", StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            // Remove the matching defined names from the workbook
            foreach (var definedName in namesToRemove)
            {
                // Names.Remove expects the name string; use the Text property which holds the name
                workbook.Worksheets.Names.Remove(definedName.Text);
            }

            // Save the modified workbook as a PDF file
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to: {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
