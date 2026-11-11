// Title: Remove named ranges that reference missing worksheets with Aspose.Cells for .NET
// Common Searches: asp.net remove named ranges pointing to deleted sheets using Aspose.Cells | c# code to clean up invalid defined names after a worksheet is removed | detect and delete Excel named ranges that reference non‑existent worksheets with Aspose.Cells | how to programmatically purge broken named ranges from a workbook in .NET | Aspose.Cells remove named ranges referencing sheets that no longer exist
// Tags: Aspose.Cells NameCollection cleanup | C# filter broken defined names | Excel invalid named range detection | remove references to deleted worksheets | programmatic named range validation .NET

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, iterates through its NameCollection, extracts each named range's RefersTo formula to identify the referenced sheet, checks whether that worksheet exists, collects indexes of names that point to missing sheets, removes those invalid named ranges in reverse order, and saves the cleaned workbook to a new file.
class RemoveInvalidNamedRanges
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of defined names (named ranges)
            NameCollection definedNames = workbook.Worksheets.Names;

            // Store indexes of names that reference missing worksheets
            List<int> indexesToRemove = new List<int>();

            for (int i = 0; i < definedNames.Count; i++)
            {
                try
                {
                    Name name = definedNames[i];
                    string refersTo = name.RefersTo; // e.g. =Sheet1!$A$1:$B$2

                    // Extract sheet name between '=' and '!'
                    int exclPos = refersTo.IndexOf('!');
                    if (exclPos > 1) // ensure a sheet name exists
                    {
                        string sheetName = refersTo.Substring(1, exclPos - 1); // skip leading '='

                        // If the worksheet does not exist, mark this name for removal
                        if (workbook.Worksheets[sheetName] == null)
                        {
                            indexesToRemove.Add(i);
                        }
                    }
                }
                catch (Exception innerEx)
                {
                    // Log and continue with next name
                    Console.WriteLine($"Warning: could not process name at index {i}. Details: {innerEx.Message}");
                }
            }

            // Remove invalid names in reverse order to keep indexes valid
            for (int i = indexesToRemove.Count - 1; i >= 0; i--)
            {
                definedNames.RemoveAt(indexesToRemove[i]);
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the cleaned workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
