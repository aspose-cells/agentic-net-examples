// Title: How to delete every slicer in an Excel workbook using Aspose.Cells for .NET (C#)
// Common Searches: C# code to remove all slicers from an Excel file with Aspose.Cells | Aspose.Cells iterate worksheets and clear slicer collections | Programmatically delete slicers from pivot tables in .xlsx using Aspose.Cells | Remove slicers from workbook before saving with Aspose.Cells .NET | How to clear worksheet slicers in Aspose.Cells C# example
// Tags: Aspose.Cells clear slicer collection | remove slicers from Excel workbook C# | iterate worksheets delete slicers Aspose | slicer removal Aspose.Cells .NET | Excel slicer cleanup using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample loads an existing Excel workbook, checks the input file, iterates through each worksheet, clears any slicers via the worksheet's Slicers collection, and saves the modified workbook, handling exceptions throughout.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and remove their slicers
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear all slicers on the current worksheet
                if (sheet.Slicers.Count > 0)
                {
                    sheet.Slicers.Clear();
                }
            }

            // Save the workbook after slicers have been removed
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
