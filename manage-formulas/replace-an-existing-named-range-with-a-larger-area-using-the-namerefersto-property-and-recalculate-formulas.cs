// Title: How to expand an existing named range and recalculate formulas with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# change named range reference to a larger area | Update Name.RefersTo property programmatically and recalculate workbook formulas | Expand MyRange to Sheet1!A1:E20 using Aspose.Cells and refresh calculations | C# code to modify named range and trigger formula recalculation in Excel file
// Tags: named range RefersTo update Aspose.Cells | expand named range area C# | trigger formula recalculation Aspose.Cells | adjust Excel named range size .NET | programmatic named range modification Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, retrieves the named range "MyRange", changes its RefersTo property to a larger range (e.g., 'Sheet1'!$A$1:$E$20), forces a full formula recalculation, and saves the updated file.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range by its name
            Name namedRange = workbook.Worksheets.Names["MyRange"];
            if (namedRange != null)
            {
                // Define the new larger area for the named range (e.g., Sheet1!$A$1:$E$20)
                string newRefersTo = "'Sheet1'!$A$1:$E$20";

                // Update the named range to refer to the new area
                namedRange.RefersTo = newRefersTo;
            }

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the updated workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific errors (e.g., password protection)
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
