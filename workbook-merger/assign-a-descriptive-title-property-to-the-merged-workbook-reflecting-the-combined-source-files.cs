// Title: Set a descriptive workbook title based on source filenames when merging Excel files with Aspose.Cells for .NET
// Common Searches: how to assign a custom title to a merged Excel workbook using Aspose.Cells C# | Aspose.Cells combine multiple .xlsx files and set built‑in document title | C# merge several workbooks and update the Title property with source file names | set BuiltInDocumentProperties.Title after copying worksheets in Aspose.Cells | create merged workbook with descriptive title from source file names in .NET
// Tags: Aspose.Cells merge multiple .xlsx workbooks | set BuiltInDocumentProperties.Title C# | assign workbook title from source filenames | copy worksheets between workbooks Aspose.Cells | update Excel document metadata Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;

// The example loads a list of Excel files, copies each worksheet into a new Workbook, gathers the base names of the source files, builds a title like "Merged Workbook: Source1, Source2, Source3", assigns this string to the workbook's BuiltInDocumentProperties.Title, and saves the merged file as MergedResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Paths of the source workbooks to be merged
            var sourceFiles = new List<string>
            {
                "Source1.xlsx",
                "Source2.xlsx",
                "Source3.xlsx"
            };

            // Create a new workbook that will contain the merged content
            var mergedWorkbook = new Workbook();

            // Remove the default empty worksheet created with a new workbook
            mergedWorkbook.Worksheets.Clear();

            // Collect the base names of source files for building the title
            var sourceNames = new List<string>();

            foreach (var filePath in sourceFiles)
            {
                // Ensure the source file exists before attempting to load it
                if (!File.Exists(filePath))
                {
                    Console.WriteLine($"Warning: File not found – {filePath}. Skipping.");
                    continue;
                }

                try
                {
                    // Load each source workbook
                    var srcWorkbook = new Workbook(filePath);
                    sourceNames.Add(Path.GetFileNameWithoutExtension(filePath));

                    // Copy every worksheet from the source workbook into the merged workbook
                    foreach (Worksheet srcSheet in srcWorkbook.Worksheets)
                    {
                        // AddCopy expects the source sheet name
                        mergedWorkbook.Worksheets.AddCopy(srcSheet.Name);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing file '{filePath}': {ex.Message}");
                }
            }

            // Construct a descriptive title that reflects the combined source files
            string descriptiveTitle = "Merged Workbook: " + string.Join(", ", sourceNames);

            // Assign the title to the workbook's built‑in properties
            mergedWorkbook.BuiltInDocumentProperties.Title = descriptiveTitle;

            // Save the merged workbook to a new file
            mergedWorkbook.Save("MergedResult.xlsx");
            Console.WriteLine("Merged workbook saved as 'MergedResult.xlsx'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
