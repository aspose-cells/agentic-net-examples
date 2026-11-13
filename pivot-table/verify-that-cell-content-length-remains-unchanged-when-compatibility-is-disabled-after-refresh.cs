// Title: Verify that a cell’s text length stays unchanged after calling Workbook.CalculateFormula with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# check if cell string length changes after CalculateFormula | how to confirm cell value length is preserved when refreshing a workbook in Aspose.Cells | C# verify cell text length remains the same after workbook recalculation Aspose.Cells | Aspose.Cells ensure cell content length is not altered by formula recalculation
// Tags: cell string length verification Aspose.Cells | calculateformula preserving cell value C# | workbook refresh cell content integrity | Aspose.Cells compatibility mode omission effect | long text cell length unchanged Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example creates a workbook, writes a long text to cell A1, records the string length, calls Workbook.CalculateFormula twice (without toggling compatibility mode), records the length again, compares the two lengths to confirm they are identical, and saves the file as Result.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put a long text into cell A1
            string longText = "This is a very long text used for testing cell content length.";
            sheet.Cells["A1"].PutValue(longText);

            // Recalculate formulas (if any) to ensure the cell content is up‑to‑date
            workbook.CalculateFormula();

            // Record the length of the cell content before any changes
            int lengthBefore = sheet.Cells["A1"].StringValue.Length;

            // (Compatibility mode toggling removed – property not available in current API)

            // Recalculate again to ensure internal updates are applied
            workbook.CalculateFormula();

            // Record the length again
            int lengthAfter = sheet.Cells["A1"].StringValue.Length;

            // Verify that the length remains unchanged
            if (lengthBefore == lengthAfter)
            {
                Console.WriteLine($"Cell content length unchanged: {lengthAfter}");
            }
            else
            {
                Console.WriteLine($"Cell content length changed from {lengthBefore} to {lengthAfter}");
            }

            // Define output file path
            string outputPath = "Result.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
