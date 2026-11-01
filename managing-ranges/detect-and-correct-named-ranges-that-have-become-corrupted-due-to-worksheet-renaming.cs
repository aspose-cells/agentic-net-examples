// Title: Fix corrupted named ranges after worksheet renaming using Aspose.Cells for .NET
// Common Searches: C# code to update Excel named ranges when a sheet is renamed with Aspose.Cells | how to programmatically repair broken defined names after worksheet rename in .NET | detect missing worksheet references in named ranges using Aspose.Cells | replace old sheet name in named range RefersTo formula with new name in C# | Aspose.Cells example for correcting named range references after sheet rename
// Tags: named range reference correction Aspose.Cells | worksheet rename mapping Aspose.Cells | validate defined name existence .NET | RefersTo property update C# | Excel named range repair Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The sample loads an Excel workbook, iterates through all defined names, parses each RefersTo formula to extract the worksheet name, checks whether the sheet still exists, and if not uses a provided old‑to‑new sheet name map to rebuild the reference. Updated named ranges are saved to a new file, restoring correct links after worksheet renaming.
class NamedRangeFixer
{
    static void Main()
    {
        // Define input and output file paths.
        string inputFilePath = "{InputFilePath}";
        string outputFilePath = "{OutputFilePath}";

        // Verify that the input file exists to avoid FileNotFoundException.
        if (!File.Exists(inputFilePath))
        {
            Console.WriteLine($"Input file not found: {inputFilePath}");
            return;
        }

        try
        {
            // Load the workbook.
            Workbook workbook = new Workbook(inputFilePath);

            // Mapping of old worksheet names to their new names.
            // Populate this dictionary with the actual renaming information.
            Dictionary<string, string> sheetRenameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
            {
                // Example: {"OldSheet", "NewSheet"},
                // Add entries as needed.
            };

            // Iterate through all defined names (named ranges) in the workbook.
            foreach (Name namedRange in workbook.Worksheets.Names)
            {
                string refersTo = namedRange.RefersTo; // e.g., =Sheet1!$A$1:$B$2
                if (string.IsNullOrEmpty(refersTo) || !refersTo.StartsWith("="))
                    continue; // Not a valid reference.

                // Strip the leading '=' for easier parsing.
                string reference = refersTo.Substring(1);
                int exclPos = reference.IndexOf('!');
                if (exclPos < 0)
                    continue; // Reference does not contain a worksheet name.

                // Extract the worksheet name part.
                string sheetName = reference.Substring(0, exclPos);
                // Remove surrounding single quotes if the sheet name contains spaces or special characters.
                if (sheetName.StartsWith("'") && sheetName.EndsWith("'"))
                    sheetName = sheetName.Substring(1, sheetName.Length - 2);

                // Verify whether the worksheet still exists.
                bool sheetExists = workbook.Worksheets[sheetName] != null;

                // If the sheet is missing and we have a rename mapping, correct the reference.
                if (!sheetExists && sheetRenameMap.TryGetValue(sheetName, out string newSheetName))
                {
                    // Rebuild the corrected reference string.
                    string correctedReference = $"={newSheetName}!{reference.Substring(exclPos + 1)}";
                    namedRange.RefersTo = correctedReference;
                    Console.WriteLine($"Corrected named range to refer to sheet '{newSheetName}'.");
                }
            }

            // Save the updated workbook.
            workbook.Save(outputFilePath);
            Console.WriteLine($"Workbook saved to: {outputFilePath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
