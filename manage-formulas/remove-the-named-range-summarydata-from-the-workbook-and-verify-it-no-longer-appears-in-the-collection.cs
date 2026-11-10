// Title: How to delete the "SummaryData" named range from an Excel workbook using Aspose.Cells for .NET and confirm its removal
// Common Searches: Aspose.Cells C# remove named range from workbook | Delete a specific named range in Excel with Aspose.Cells .NET | Check if a named range still exists after removal using Aspose.Cells | Workbook.Worksheets.Names.Remove example in C# code | Verify named range deletion in an Aspose.Cells workbook
// Tags: Workbook.Names.Remove method Aspose.Cells | delete Excel named range programmatically | named range removal verification Aspose.Cells | C# Aspose.Cells named range management | remove specific named range SummaryData

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing Excel file, retrieves the named range "SummaryData", removes it with Workbook.Worksheets.Names.Remove, verifies that the range no longer appears in the Names collection, and saves the updated workbook, handling any exceptions that may arise.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";
            const string rangeName = "SummaryData";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Attempt to retrieve the named range
            Name namedRange = workbook.Worksheets.Names[rangeName];

            // Remove the named range if it exists
            if (namedRange != null)
            {
                // Correct API: remove by name string
                workbook.Worksheets.Names.Remove(rangeName);
                Console.WriteLine($"Named range \"{rangeName}\" removed.");
            }
            else
            {
                Console.WriteLine($"Named range \"{rangeName}\" does not exist.");
            }

            // Verify that the named range no longer exists
            bool stillExists = workbook.Worksheets.Names[rangeName] != null;
            Console.WriteLine($"SummaryData exists after removal: {stillExists}");

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
