// Title: Delete the named range "ObsoleteRange" from an Excel workbook using Aspose.Cells for .NET and verify that it no longer exists
// AI Prompts: Load an Excel file with Aspose.Cells, remove the named range "ObsoleteRange", and save the workbook to a new file. | After deleting a named range, query the workbook to confirm the range is absent and output a success or failure message.
// Common Searches: aspocells c# delete specific named range from workbook | how to confirm a named range was removed using Aspose.Cells | remove named range and save workbook with Aspose.Cells .NET | check if a named range exists after deletion in C# Aspose.Cells
// Tags: Aspose.Cells remove named range | C# verify named range deletion | Aspose.Cells workbook save after range removal | C# Excel named range existence check

using Aspose.Cells;
using System;
using System.IO;

// The example loads "input.xlsx", deletes the named range "ObsoleteRange" if present, confirms the range is no longer in the workbook, and saves the result as "output.xlsx" while handling missing files and exceptions.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "ObsoleteRange";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve the named range
            Name namedRange = workbook.Worksheets.Names[rangeName];

            // Delete the named range if it exists
            if (namedRange != null)
            {
                // Remove by name (correct API overload)
                workbook.Worksheets.Names.Remove(rangeName);
                Console.WriteLine($"Named range \"{rangeName}\" removed.");
            }
            else
            {
                Console.WriteLine($"Named range \"{rangeName}\" does not exist.");
            }

            // Verify removal
            bool stillExists = workbook.Worksheets.Names[rangeName] != null;
            Console.WriteLine(stillExists
                ? "Named range still exists."
                : "Named range successfully removed.");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
