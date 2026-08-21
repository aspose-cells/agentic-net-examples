// Title: C# – Count Cells with Custom Number Formats per Worksheet using Aspose.Cells
// Description: Loads a workbook, iterates through each worksheet, enumerates instantiated cells, checks the Style.Custom property, tallies cells that use custom number formats, and prints the count for every sheet. Optionally saves the workbook.
// Keywords: Aspose.Cells custom number format count | C# enumerate cells Aspose.Cells | worksheet custom format statistics | detect custom number formats .NET | Aspose.Cells style.Custom property
// Common Searches: count cells with custom number formats Aspose.Cells C# | how to list worksheets with custom formatted cells | enumerate instantiated cells and check custom format | Aspose.Cells report custom number format usage per sheet
// Developer Intent: Determine how many cells on each worksheet use a custom number format.
// Use Cases: Audit a workbook to see the prevalence of custom number formats before distribution. | Validate compliance by ensuring custom formats stay within a defined limit. | Create a summary sheet that lists each worksheet alongside its custom‑format cell count.
// AI Prompts: Generate C# code with Aspose.Cells that counts custom‑format cells per worksheet and writes the results to a new summary worksheet. | Explain how to modify the sample to also capture the addresses of cells that use custom number formats. | Provide guidance on excluding built‑in formats so only truly custom number formats are counted in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

// Loads a workbook, iterates through each worksheet, enumerates instantiated cells, checks the Style.Custom property, tallies cells that use custom number formats, and prints the count for every sheet. Optionally saves the workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        for (int wsIndex = 0; wsIndex < workbook.Worksheets.Count; wsIndex++)
        {
            Worksheet worksheet = workbook.Worksheets[wsIndex];
            Cells cells = worksheet.Cells;
            int customFormatCount = 0;

            // Enumerate all instantiated cells in the worksheet
            foreach (Cell cell in cells)
            {
                // Retrieve the cell's style
                Style style = cell.GetStyle();

                // If the Custom property is not empty, the cell uses a custom number format
                if (!string.IsNullOrEmpty(style.Custom))
                {
                    customFormatCount++;
                }
            }

            // Report the count for the current worksheet
            Console.WriteLine($"Worksheet \"{worksheet.Name}\" contains {customFormatCount} cells with custom number formats.");
        }

        // Save the workbook (optional, adjust path as needed)
        workbook.Save("output.xlsx");
    }
}
