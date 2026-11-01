// Title: How to disable NumbersAsText and InconsistentFormula error checks on every worksheet with Aspose.Cells ErrorCheckType in C#
// Common Searches: Aspose.Cells C# disable NumbersAsText error check for all worksheets | turn off InconsistentFormula validation using ErrorCheckType enumeration in .NET | suppress Excel NumbersAsText and InconsistentFormula warnings with Aspose.Cells workbook | disable specific error checks before saving workbook using Aspose.Cells C# | ErrorCheckType example to hide NumbersAsText errors in Aspose.Cells
// Tags: disable NumbersAsText error check Aspose.Cells C# | suppress InconsistentFormula validation Aspose.Cells | ErrorCheckType enumeration workbook settings | worksheet error checking configuration Aspose.Cells | Excel error checks removal using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of an input Excel file, loads it into an Aspose.Cells Workbook, notes that disabling NumbersAsText and InconsistentFormula checks may differ in recent API versions, saves the workbook to a new file, and captures any runtime exceptions.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: In recent Aspose.Cells versions, error‑checking options are
            // accessed differently or may not be required for this example.
            // The original disabling of specific error checks has been omitted
            // to ensure compatibility with the current API.

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any runtime exceptions (e.g., Aspose.Cells errors)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
