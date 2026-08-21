// Title: Aspose.Cells for .NET – Clear the contents of the "ReportData" named range while keeping its definition
// Description: Loads an Excel workbook, locates the named range "ReportData", uses the Range.ClearContents() method to erase only cell values, and saves the file. The named range itself remains intact, and the code safely handles a missing range.
// Keywords: Aspose.Cells clear named range | C# clear range contents | preserve named range definition | Excel named range Aspose.Cells | ClearContents method .NET
// Common Searches: Aspose.Cells how to clear values of a named range | C# remove data from ReportData range without deleting name | Clear cells in a specific named range using Aspose.Cells | Keep named range after clearing its contents in .NET
// Developer Intent: Remove all cell values from the "ReportData" named range without deleting the range itself.
// Use Cases: Reset a reporting section before inserting fresh data. | Empty a template area while preserving formulas that reference the named range. | Prepare a workbook for user input by clearing previous entries but keeping the range for later reuse.
// AI Prompts: Write C# code with Aspose.Cells that clears only the contents of a named range called "ReportData" and saves the workbook. | Show how to retrieve a named range, call ClearContents, and retain the range definition using Aspose.Cells for .NET. | Explain error handling when the named range "ReportData" does not exist while attempting to clear its contents with Aspose.Cells.

using Aspose.Cells;
using System;
using System.IO;

// Loads an Excel workbook, locates the named range "ReportData", uses the Range.ClearContents() method to erase only cell values, and saves the file. The named range itself remains intact, and the code safely handles a missing range.
class ClearNamedRange
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists; create a placeholder workbook if missing
            if (!File.Exists(inputPath))
            {
                var placeholder = new Workbook();
                placeholder.Worksheets[0].Name = "Sheet1";
                placeholder.Save(inputPath);
            }

            // Load the workbook that contains the named range
            Workbook workbook = new Workbook(inputPath);

            // Retrieve the named range "ReportData"
            Name namedRange = workbook.Worksheets.Names["ReportData"];
            if (namedRange != null)
            {
                // Obtain the Range object linked to the name
                Aspose.Cells.Range range = namedRange.GetRange();

                // Clear only the cell contents; the range definition remains intact
                range.ClearContents();
            }

            // Save the modified workbook
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
