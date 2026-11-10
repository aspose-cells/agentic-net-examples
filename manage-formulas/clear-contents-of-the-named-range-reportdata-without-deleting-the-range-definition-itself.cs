// Title: Clear the contents of a named range while preserving its definition using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clear values in a named range without removing the range | How to use GetRangeByName to clear a specific range in an Excel workbook with Aspose.Cells | Remove formulas from a named range but keep the range definition in .NET | Preserve Excel named range definition while clearing its data using Aspose.Cells | Clear contents of ReportData range and save workbook with Aspose.Cells
// Tags: ClearContents method Aspose.Cells | named range manipulation C# | GetRangeByName usage Aspose.Cells | clear Excel range values without deleting definition | Aspose.Cells workbook save after clearing range

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel file, retrieves the named range "ReportData" with GetRangeByName, calls ClearContents() to erase values and formulas while keeping the named range definition, and then saves the workbook. It includes file existence checks and basic error handling.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ReportData"
            // Use fully qualified type to avoid ambiguity with System.Range
            Aspose.Cells.Range reportRange = workbook.Worksheets.GetRangeByName("ReportData");

            if (reportRange != null)
            {
                // Clear only the contents (values and formulas) of the range.
                // The named range definition itself remains intact.
                reportRange.ClearContents();
            }
            else
            {
                Console.WriteLine("Named range 'ReportData' not found.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
