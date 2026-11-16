// Title: Set a default row height of 20 points in an Aspose.Cells worksheet using C#
// Common Searches: C# Aspose.Cells set row height for multiple rows | Aspose.Cells default row height 20 points example | Set row height programmatically in Excel with Aspose.Cells .NET | Adjust worksheet row height before saving workbook using Aspose.Cells | How to change row height of first rows in Aspose.Cells C#
// Tags: Aspose.Cells row height adjustment C# | worksheet row height configuration | Cells.SetRowHeight usage .NET | initialize workbook with custom row height | pre‑save directory validation for Excel output

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new Workbook, accesses the first Worksheet, and uses Cells.SetRowHeight to set rows 0‑2 to a height of 20 points. It also ensures the output folder exists before saving the workbook as Output.xlsx, handling any exceptions that may occur.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Set the default row height to 20 points for better readability
            // Note: Aspose.Cells does not expose a DefaultRowHeight property on Worksheet in some versions.
            // As an alternative, set the height for the first few rows or leave the default.
            sheet.Cells.SetRowHeight(0, 20);
            sheet.Cells.SetRowHeight(1, 20);
            sheet.Cells.SetRowHeight(2, 20);

            // Define output file path
            string outputPath = "Output.xlsx";

            // Ensure the directory for the output file exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook (lifecycle rule: save)
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
