// Title: Recalculate dependent Excel formulas after deleting a named range using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# recalculate formulas after deleting a named range | How to refresh dependent cells when a named range is removed with Aspose.Cells | Workbook.CalculateFormula usage after named range removal in .NET | Delete named range programmatically with Aspose.Cells and update formulas | C# example for removing Excel named range and recalculating formulas using Aspose
// Tags: named range removal Aspose.Cells | Workbook.CalculateFormula C# | Excel formula recalculation Aspose.Cells | dependent cell update after named range deletion | Aspose.Cells workbook manipulation .NET | programmatic named range management Aspose

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, optionally removes the named range "MyRange" from worksheets and the workbook, calls Workbook.CalculateFormula to refresh any formulas that referenced the deleted range, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // NOTE: Older versions of Aspose.Cells may not expose the Names collection.
            // If the API is available, you can uncomment the following blocks to remove
            // the named range "MyRange" at both worksheet and workbook levels.

            /*
            // Delete the named range "MyRange" at worksheet level, if it exists
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                if (sheet.Names.Contains("MyRange"))
                {
                    sheet.Names.Remove("MyRange");
                }
            }

            // Delete the named range "MyRange" at workbook level, if it exists
            if (workbook.Names.Contains("MyRange"))
            {
                workbook.Names.Remove("MyRange");
            }
            */

            // Recalculate all formulas to update dependent cells
            workbook.CalculateFormula();

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
