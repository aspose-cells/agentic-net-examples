// Title: Set Worksheet Paper Size from a Configuration Map While Loading a Workbook – Aspose.Cells for .NET
// Description: This C# example shows how to load an Excel file with Aspose.Cells, read a dictionary that maps worksheet names to PaperSizeType values, assign the corresponding page‑setup paper size to each matching sheet, and save the updated workbook.
// Keywords: Aspose.Cells set worksheet paper size | C# paper size mapping Excel | load workbook change page setup | dictionary based page setup Aspose | configure print settings programmatically | PaperSizeType example | Aspose.Cells LoadOptions
// Common Searches: how to set different paper sizes for each sheet using Aspose.Cells .NET | apply worksheet name to paper size mapping when loading Excel | C# change page setup paper size per worksheet | Aspose.Cells assign A4 to Sheet1 and Letter to Sheet2 | programmatically update print settings in Excel files
// Developer Intent: Apply a predefined name‑to‑paper‑size mapping to worksheets during workbook load.
// Use Cases: Load a template and automatically set A4 for Sheet1, Letter for Sheet2, A3 for Report before distribution. | Generate reports where each worksheet requires a specific print format without manual editing. | Create a utility that reads a JSON or XML config and updates page‑setup settings of existing workbooks.
// AI Prompts: Write C# code with Aspose.Cells that reads a JSON file containing worksheet‑name to PaperSizeType mappings and applies them during workbook load. | Show how to log worksheet names that are missing from the configuration and skip them gracefully. | Demonstrate using LoadOptions to keep existing page‑setup options while overriding only the paper size for selected sheets.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsPaperSizeMapping
{
    // This C# example shows how to load an Excel file with Aspose.Cells, read a dictionary that maps worksheet names to PaperSizeType values, assign the corresponding page‑setup paper size to each matching sheet, and save the updated workbook.
    class Program
    {
        static void Main()
        {
            // Configuration mapping: worksheet name -> desired paper size
            var sheetPaperSizes = new Dictionary<string, PaperSizeType>(StringComparer.OrdinalIgnoreCase)
            {
                { "Sheet1", PaperSizeType.PaperA4 },
                { "Sheet2", PaperSizeType.PaperLetter },
                { "Report", PaperSizeType.PaperA3 }
                // Add more mappings as needed
            };

            // Create load options (default settings)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook with the specified options
            Workbook workbook = new Workbook("input.xlsx", loadOptions);

            // Apply paper size to each worksheet based on the configuration mapping
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (sheetPaperSizes.TryGetValue(ws.Name, out PaperSizeType size))
                {
                    // Set the paper size for the current worksheet
                    ws.PageSetup.PaperSize = size;
                }
            }

            // Save the modified workbook
            workbook.Save("output.xlsx", SaveFormat.Xlsx);
        }
    }
}
