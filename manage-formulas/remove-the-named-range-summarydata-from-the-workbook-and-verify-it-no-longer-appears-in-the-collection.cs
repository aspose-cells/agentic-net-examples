// Title: Delete the 'SummaryData' named range from an Excel workbook using Aspose.Cells for .NET and verify its removal
// Common Searches: Aspose.Cells C# remove a specific named range from a workbook | how to delete the SummaryData named range with Aspose.Cells | check if a named range still exists after removal in Aspose.Cells | C# code to remove and confirm deletion of an Excel named range using Aspose.Cells | remove named range and save workbook with Aspose.Cells .NET
// Tags: Aspose.Cells named range removal | C# Excel named range deletion | Workbook.Names collection manipulation Aspose | verify named range absence Aspose.Cells | save workbook after named range removal C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells, accesses the workbook's named ranges collection, removes the "SummaryData" named range if present, confirms that it no longer exists, and saves the updated workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Access the collection of named ranges
            var namedRanges = workbook.Worksheets.Names;

            // Remove the named range "SummaryData" if it exists
            if (namedRanges["SummaryData"] != null)
            {
                namedRanges.Remove("SummaryData");
            }

            // Verify that "SummaryData" no longer exists
            bool stillExists = namedRanges["SummaryData"] != null;
            Console.WriteLine($"SummaryData exists after removal: {stillExists}");

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
