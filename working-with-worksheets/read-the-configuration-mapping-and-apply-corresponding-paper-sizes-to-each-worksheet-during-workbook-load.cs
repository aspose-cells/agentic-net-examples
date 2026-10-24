// Title: Apply worksheet-specific paper sizes from a dictionary when loading an Excel workbook with Aspose.Cells for .NET
// Common Searches: how to assign different paper sizes to worksheets using Aspose.Cells in C# | using a dictionary to map worksheet names to PaperSizeType with Aspose.Cells | set default A4 paper size for worksheets not in mapping Aspose.Cells | load an Excel file and change page setup paper size per sheet Aspose.Cells .NET | error handling when loading and saving workbook after changing paper size Aspose.Cells
// Tags: worksheet page setup paper size Aspose.Cells | dictionary based PaperSizeType mapping | default paper size for unmapped sheets | load workbook modify page setup .NET | apply custom paper size during workbook load

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example creates a case‑insensitive dictionary that maps worksheet names to Aspose.Cells PaperSizeType values, loads an Excel workbook, iterates through each worksheet to set the mapped paper size or fall back to A4, ensures the output directory exists, and saves the workbook while handling file‑related errors.
class Program
{
    static void Main()
    {
        try
        {
            // Mapping of worksheet names to desired paper sizes
            var paperSizeMapping = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", PaperSizeType.PaperA4 },
                { "Report", PaperSizeType.PaperLetter },
                { "Summary", PaperSizeType.PaperA3 }
            };

            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from file
            Workbook workbook;
            try
            {
                workbook = new Workbook(inputPath);
            }
            catch (Exception loadEx)
            {
                Console.WriteLine($"Failed to load workbook: {loadEx.Message}");
                return;
            }

            // Apply the configured paper size to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMapping.TryGetValue(sheet.Name, out PaperSizeType size))
                {
                    sheet.PageSetup.PaperSize = size;
                }
                else
                {
                    // Use a default paper size when no mapping exists
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
