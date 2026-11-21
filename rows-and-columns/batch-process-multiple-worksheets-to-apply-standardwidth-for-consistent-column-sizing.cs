// Title: Set a consistent column width for every worksheet in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set same column width for all sheets in a workbook | How to apply a standard column width to every worksheet with Aspose.Cells | Batch update column width across multiple worksheets using Aspose.Cells .NET | C# code to set Cells.StandardWidth for all worksheets in an Excel file | Automate uniform column sizing in Excel workbooks with Aspose.Cells
// Tags: set Cells.StandardWidth Aspose.Cells | uniform column width across worksheets | batch column width update Excel .NET | standard column width workbook automation | Aspose.Cells column sizing for multiple sheets

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing .xlsx file, iterates through all worksheets, assigns a fixed Cells.StandardWidth (e.g., 15 characters) to each sheet, and saves the workbook, including checks for the input file and basic exception handling.
class Program
{
    static void Main()
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook from the existing file
            Workbook workbook = new Workbook(inputPath);

            // Define the standard column width (in characters) to apply to all worksheets
            double standardWidth = 15.0;

            // Apply the standard width to each worksheet via its Cells collection
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                sheet.Cells.StandardWidth = standardWidth;
            }

            // Save the updated workbook to a new file
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, Aspose.Cells specific issues)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
