// Title: How to delete the named range "ObsoleteRange" from an Excel workbook and confirm its removal using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete specific named range and verify it no longer exists | remove a named range from an .xlsx file programmatically with Aspose.Cells | check if a named range was successfully deleted before saving workbook in .NET | C# code to find and delete "ObsoleteRange" using Aspose.Cells API | ensure named range removal after modifying an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells delete named range | C# remove Excel named range | named range existence verification Aspose.Cells | Aspose.Cells workbook save after range removal | programmatic Excel named range management .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file with Aspose.Cells, searches for the named range "ObsoleteRange", removes it if present, iterates the workbook's names collection to confirm the range is gone, creates the output folder if needed, saves the modified workbook, and logs each step to the console.
class DeleteNamedRangeExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify input file existence
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load workbook
            Workbook workbook = new Workbook(inputPath);

            // Find the named range "ObsoleteRange"
            Name? targetRange = null;
            foreach (Name n in workbook.Worksheets.Names)
            {
                if (string.Equals(n.Text, "ObsoleteRange", StringComparison.OrdinalIgnoreCase))
                {
                    targetRange = n;
                    break;
                }
            }

            // Delete the named range if it exists
            if (targetRange != null)
            {
                workbook.Worksheets.Names.Remove(targetRange.Text);
                Console.WriteLine("ObsoleteRange has been removed.");
            }
            else
            {
                Console.WriteLine("ObsoleteRange not found; nothing to delete.");
            }

            // Verify removal
            bool exists = false;
            foreach (Name n in workbook.Worksheets.Names)
            {
                if (string.Equals(n.Text, "ObsoleteRange", StringComparison.OrdinalIgnoreCase))
                {
                    exists = true;
                    break;
                }
            }
            Console.WriteLine($"ObsoleteRange exists after deletion: {exists}");

            // Ensure output directory exists
            string? outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
