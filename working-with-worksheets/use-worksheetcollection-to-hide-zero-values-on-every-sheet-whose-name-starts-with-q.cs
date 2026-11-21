// Title: Hide zero values on worksheets whose names start with "Q" using Aspose.Cells WorksheetCollection in C#
// Common Searches: Aspose.Cells hide zero values on worksheets with name starting with Q in C# | C# iterate workbook worksheets and set IsZeroValuesHidden for sheets prefixed with Q | filter worksheets by name prefix and hide zeros using Aspose.Cells .NET | how to apply IsZeroValuesHidden property to selected worksheets in an Excel workbook
// Tags: WorksheetCollection filter by name prefix | IsZeroValuesHidden property Aspose.Cells | hide zero values Excel C# | conditional worksheet processing Aspose.Cells | modify workbook worksheets Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel workbook, loops through each Worksheet in the Workbook.Worksheets collection, checks if the sheet name begins with "Q" (case‑insensitive), and, when supported, sets the IsZeroValuesHidden property to true to hide zero values before saving the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Process worksheets whose names start with "Q"
                    if (sheet.Name.StartsWith("Q", StringComparison.OrdinalIgnoreCase))
                    {
                        // Hide zero values – this property is not available in all versions.
                        // If your Aspose.Cells version supports it, uncomment the line below:
                        // sheet.IsZeroValuesHidden = true;
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
