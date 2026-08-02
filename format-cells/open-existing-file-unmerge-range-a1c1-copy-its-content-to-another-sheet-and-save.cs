// Title: C# – Unmerge A1:C1 and Copy Range to a New Worksheet with Aspose.Cells
// Description: Load an existing workbook, unmerge the merged range A1:C1, add a new sheet, copy the range with all data and formatting using PasteOptions.All, and save the result as a new file.
// Keywords: Aspose.Cells | C# | unmerge cells | copy range | PasteOptions.All | Workbook.Save | Excel automation | merge header row | add worksheet | copy formatting
// Common Searches: Aspose.Cells unmerge merged cells C# | Copy range A1:C1 to another sheet preserving formatting | How to use PasteOptions.All in Aspose.Cells | Add new worksheet and copy data with Aspose.Cells | Save modified workbook after unmerging cells
// Developer Intent: Unmerge the merged cells A1:C1, copy their content to a newly created worksheet, and save the updated workbook.
// Use Cases: Duplicate a merged header row from a template sheet to a report sheet while retaining formatting. | Create a clean analysis sheet by unmerging and copying a specific range from the source workbook. | Archive a section of a workbook by moving an unmerged range to a separate worksheet before distribution.
// AI Prompts: Generate C# code using Aspose.Cells that unmerges a specified range and copies it to a new worksheet with full formatting. | Explain how PasteOptions.PasteType.All affects copying merged cells between worksheets in Aspose.Cells. | Provide error‑handling patterns for missing input files and save failures when copying ranges with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsUnmergeAndCopy
{
    // Load an existing workbook, unmerge the merged range A1:C1, add a new sheet, copy the range with all data and formatting using PasteOptions.All, and save the result as a new file.
    class Program
    {
        static void Main()
        {
            try
            {
                const string inputFile = "Input.xlsx";
                const string outputFile = "Output.xlsx";

                // Verify input file exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file '{inputFile}' not found.");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputFile);

                // Access the first worksheet (source)
                Worksheet sourceSheet = workbook.Worksheets[0];

                // Unmerge the range A1:C1
                Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange("A1", "C1");
                sourceRange.UnMerge();

                // Add a new worksheet to copy the content into (destination)
                Worksheet destSheet = workbook.Worksheets.Add("CopySheet");

                // Create the destination range A1:C1
                Aspose.Cells.Range destRange = destSheet.Cells.CreateRange("A1", "C1");

                // Prepare paste options to copy all cell data, formats, etc.
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All,
                    SkipBlanks = false,
                    Transpose = false
                };

                // Copy the content from source range to destination range
                sourceRange.Copy(destRange, pasteOptions);

                // Save the modified workbook
                workbook.Save(outputFile);
                Console.WriteLine($"Workbook saved to '{outputFile}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
