// Title: Delete the named range "ObsoleteRange" from an Excel workbook using Aspose.Cells for .NET and confirm its removal
// Common Searches: Aspose.Cells C# remove named range ObsoleteRange from workbook | C# verify that a named range has been deleted with Aspose.Cells | How to delete a specific named range and check its existence using Aspose.Cells .NET | Remove a named range by name and save the workbook with Aspose.Cells | Aspose.Cells .NET confirm named range no longer present after deletion
// Tags: delete named range Aspose.Cells .NET | named range removal Excel Aspose.Cells | verify named range absence C# | Aspose.Cells workbook save after range deletion | remove range by name Aspose.Cells API

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, accesses the named range "ObsoleteRange" via the workbook's Names collection, removes it if present, checks that the range no longer exists, and saves the modified workbook to a new file, handling missing input files and creating output directories as needed.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string targetRangeName = "ObsoleteRange";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range object by its name using the indexer
            Name namedRange = workbook.Worksheets.Names[targetRangeName];

            // Delete the named range if it exists
            if (namedRange != null)
            {
                // Remove by name (string) rather than Name object
                workbook.Worksheets.Names.Remove(targetRangeName);
            }

            // Verify that the named range has been removed
            bool stillExists = workbook.Worksheets.Names[targetRangeName] != null;
            Console.WriteLine($"ObsoleteRange exists after deletion: {stillExists}");

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
