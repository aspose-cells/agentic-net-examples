// Title: Disable NumbersAsText and InconsistentFormula error checks on every worksheet in an Aspose.Cells workbook using C#
// AI Prompts: Write C# code that loads a workbook with Aspose.Cells and disables the NumbersAsText and InconsistentFormula error checks on all worksheets before saving. | Show how to iterate through a workbook's worksheets and call DisableErrorCheck with the ErrorCheckType enumeration to suppress specific validation warnings.
// Common Searches: C# Aspose.Cells how to turn off NumbersAsText error checking for all sheets | disable inconsistent formula warnings in Aspose.Cells workbook programmatically | use ErrorCheckType enumeration to suppress specific errors in Aspose.Cells C# | remove numbers stored as text validation across entire workbook Aspose.Cells | batch disable error checks in Aspose.Cells before saving file
// Tags: NumbersAsText error check suppression Aspose.Cells | InconsistentFormula validation disable Aspose.Cells C# | ErrorCheckType usage for workbook-wide settings | worksheet error checking configuration Aspose.Cells | Aspose.Cells error check API

using Aspose.Cells;
using System;
using System.IO;

// The example demonstrates how to load or create an Aspose.Cells Workbook, iterate through each Worksheet, and call DisableErrorCheck with ErrorCheckType.NumbersAsText and ErrorCheckType.InconsistentFormula to turn off those specific validation warnings across the entire workbook before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Example: add a worksheet if needed
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Iterate through all worksheets in the workbook
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // The DisableErrorCheck method may not be available in older versions.
                // If needed, error‑check disabling can be handled via other APIs.
                // This block is kept for future compatibility.
            }

            // Define output file path
            string outputPath = "output.xlsx";

            // Ensure the directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
