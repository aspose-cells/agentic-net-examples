// Title: Apply worksheet paper sizes from a configuration dictionary during workbook load – Aspose.Cells for .NET
// Description: Demonstrates how to load an Excel file with Aspose.Cells, use a case‑insensitive Dictionary<string, PaperSizeType> to map worksheet names to paper sizes, set each sheet's PageSetup.PaperSize, and save the updated workbook.
// Keywords: Aspose.Cells C# set worksheet paper size | paper size mapping Aspose.Cells | LoadOptions workbook page setup | Dictionary worksheet name to PaperSizeType | Excel print settings Aspose.Cells
// Common Searches: Aspose.Cells set different paper sizes per sheet | C# map worksheet names to paper sizes | Load workbook and change page setup with Aspose.Cells | Apply configuration file to Excel sheet print settings | How to use PaperSizeType enum in Aspose.Cells
// Developer Intent: Assign the appropriate PaperSize to each worksheet based on a predefined name‑to‑size map when opening an Excel workbook.
// Use Cases: Automatically apply A4, Letter, or A3 sizes to specific sheets after loading a template workbook. | Read a JSON or XML file into a Dictionary<string, PaperSizeType> and enforce those print settings across all worksheets. | Create new workbooks programmatically and ensure each sheet prints on the correct paper size before distribution.
// AI Prompts: Generate C# code that reads a JSON file containing worksheet names and PaperSizeType values, builds the mapping dictionary, and applies the sizes to a loaded Aspose.Cells workbook. | Show how to log worksheet names that are missing from the mapping dictionary while continuing the paper‑size assignment. | Refactor the example into a reusable method ApplyPaperSizeMapping(Workbook workbook, IDictionary<string, PaperSizeType> mapping).

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeMapping
{
    // Demonstrates how to load an Excel file with Aspose.Cells, use a case‑insensitive Dictionary<string, PaperSizeType> to map worksheet names to paper sizes, set each sheet's PageSetup.PaperSize, and save the updated workbook.
    class Program
    {
        static void Main()
        {
            // Configuration mapping: worksheet name -> desired paper size
            var paperSizeMapping = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", PaperSizeType.PaperA4 },
                { "Sheet2", PaperSizeType.PaperLetter },
                { "Report", PaperSizeType.PaperA3 }
            };

            // Create LoadOptions (default settings)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("InputWorkbook.xlsx", loadOptions);

            // Apply paper size to each worksheet based on the mapping
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (paperSizeMapping.TryGetValue(sheet.Name, out PaperSizeType size))
                {
                    // Set the paper size for the current worksheet
                    sheet.PageSetup.PaperSize = size;
                }
            }

            // Save the modified workbook
            workbook.Save("OutputWorkbook.xlsx", SaveFormat.Xlsx);
        }
    }
}
