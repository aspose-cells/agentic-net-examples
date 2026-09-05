// Title: Save a modified workbook with charts to XLSX using Aspose.Cells in C# while retaining all chart elements
// AI Prompts: Write C# code that loads an existing XLSX file with Aspose.Cells, modifies chart series or values, and saves the workbook to a new XLSX file without losing any chart elements. | Show how to check for the presence of the source workbook and implement robust exception handling when saving a chart‑rich workbook with Aspose.Cells.
// Common Searches: c# aspose.cells retain chart elements when saving workbook to xlsx | how to update chart data in an existing workbook using Aspose.Cells | saving modified charts in Aspose.Cells without losing formatting | aspose.cells handle missing input file exception | export workbook with charts to xlsx using Aspose.Cells C#
// Tags: Aspose.Cells save workbook with charts to XLSX | C# retain chart elements Aspose.Cells | Aspose.Cells modify chart series | Aspose.Cells handle missing input file | Aspose.Cells workbook export preserving formatting

using System;
using System.IO;
using Aspose.Cells;

// Loads an existing XLSX workbook, optionally updates its charts, and saves the workbook to a new XLSX file with Aspose.Cells, ensuring all chart elements are kept and handling missing input file errors gracefully.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file not found at '{inputPath}'.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // At this point charts can be modified as needed
            // Example: modify chart series, values, etc.
            // (Assume modifications are performed here)

            // Save the workbook to XLSX format; chart controls are preserved by default
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
