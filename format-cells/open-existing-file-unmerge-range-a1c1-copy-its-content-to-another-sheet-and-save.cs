// Title: C# – Unmerge A1:C1 and copy to a new worksheet using Aspose.Cells
// Description: Loads an existing workbook, unmerges the merged range A1:C1 on the first sheet, adds a worksheet named "CopySheet", copies the unmerged content (values, formulas, formatting) to the same range on the new sheet with PasteOptions.All, and saves the result as output.xlsx.
// Keywords: Aspose.Cells | C# unmerge cells | copy range to another sheet | PasteOptions.All | .NET spreadsheet manipulation | unmerge A1:C1 | duplicate merged header | save workbook Aspose
// Common Searches: Aspose.Cells unmerge merged cells C# | copy range A1:C1 to new worksheet Aspose.Cells | how to duplicate merged header row with Aspose.Cells .NET | C# unmerge and copy cells between worksheets | Aspose.Cells paste options example
// Developer Intent: Unmerge the merged range A1:C1, copy its full content to a newly created worksheet, and save the updated workbook.
// Use Cases: Extract a merged title row from a source sheet and place it on a separate reporting sheet. | Create a template copy of a merged header while preserving styles, formulas, and values. | Prepare a workbook for downstream processing by isolating merged header cells on their own sheet.
// AI Prompts: Write C# code that uses Aspose.Cells to unmerge range A1:C1, copy it to a new worksheet called "CopySheet", and save the file as output.xlsx. | Show how to copy a merged range with all formatting, formulas, and values using PasteOptions.All in Aspose.Cells for .NET. | Explain step‑by‑step how to handle merged cells, duplicate them on another sheet, and persist the workbook with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // Loads an existing workbook, unmerges the merged range A1:C1 on the first sheet, adds a worksheet named "CopySheet", copies the unmerged content (values, formulas, formatting) to the same range on the new sheet with PasteOptions.All, and saves the result as output.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputPath = "input.xlsx";
                const string outputPath = "output.xlsx";

                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Source worksheet (first sheet)
                Worksheet srcSheet = workbook.Worksheets[0];

                // Unmerge the range A1:C1
                srcSheet.Cells.UnMerge(0, 0, 1, 3);

                // Add a new worksheet to copy the content into
                Worksheet destSheet = workbook.Worksheets.Add("CopySheet");

                // Define source and destination ranges (A1:C1)
                AsposeRange srcRange = srcSheet.Cells.CreateRange("A1", "C1");
                AsposeRange destRange = destSheet.Cells.CreateRange("A1", "C1");

                // Set paste options (copy all content)
                PasteOptions options = new PasteOptions
                {
                    PasteType = PasteType.All,
                    SkipBlanks = false,
                    Transpose = false
                };

                // Copy the source range to the destination range
                srcRange.Copy(destRange, options);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
