// Title: Set worksheet-specific paper sizes from a name-to-PaperSizeType map when loading an Excel workbook with Aspose.Cells for .NET
// Common Searches: how to assign different paper sizes to worksheets using Aspose.Cells in C# | Aspose.Cells set page setup paper size based on worksheet name | load Excel file and apply custom paper size per sheet with Aspose.Cells .NET | C# dictionary mapping worksheet names to PaperSizeType for Aspose.Cells | default paper size fallback when worksheet not in mapping Aspose.Cells
// Tags: Aspose.Cells worksheet paper size configuration | C# sheet name to PaperSizeType dictionary | default A4 fallback Aspose.Cells page setup | load workbook apply per-sheet page setup | Aspose.Cells .NET map worksheet names to paper sizes

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, uses a case‑insensitive C# dictionary to map worksheet names to Aspose.Cells PaperSizeType values, sets each sheet's PageSetup.PaperSize (defaulting to A4 when no mapping exists), and saves the modified workbook.
class ApplyPaperSizes
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

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

            // Mapping: worksheet name -> desired paper size (using PaperSizeType for compatibility)
            var paperSizeMapping = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", PaperSizeType.PaperA4 },
                { "Report", PaperSizeType.PaperLetter },
                { "Data",   PaperSizeType.PaperLegal }
                // Add more mappings as needed
            };

            // Apply the paper size to each worksheet based on the mapping
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMapping.TryGetValue(sheet.Name, out PaperSizeType size))
                {
                    // Set the paper size using the PageSetup property
                    sheet.PageSetup.PaperSize = size;
                }
                else
                {
                    // Optional: set a default paper size if no mapping is found
                    sheet.PageSetup.PaperSize = PaperSizeType.PaperA4;
                }
            }

            // Path to the output workbook
            string outputPath = "output.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime exceptions gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
