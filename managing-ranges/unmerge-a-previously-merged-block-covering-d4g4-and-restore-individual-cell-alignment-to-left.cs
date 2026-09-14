// Title: How to unmerge cells D4:G4 and set each cell’s alignment to left using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, removes the merged block spanning D4 to G4, and aligns each resulting cell to the left. | Write a .NET snippet that creates a Range for D4:G4, calls UnMerge, then loops through the cells to set Style.HorizontalAlignment = TextAlignmentType.Left.
// Common Searches: Aspose.Cells how to split merged cells D4:G4 and keep left alignment | C# unmerge specific Excel range and set each cell's alignment using Aspose.Cells | example code for unmerging a block and applying left alignment in Aspose.Cells .NET | remove merged cells D4 to G4 and adjust cell style with Aspose.Cells | Aspose.Cells unmerge range then apply horizontal alignment left
// Tags: Aspose.Cells unmerge specific range | apply left horizontal alignment Aspose.Cells | C# iterate cells in unmerged range | Excel .xlsx range style update Aspose.Cells | modify cell style after unmerge .NET

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the presence of input.xlsx, loads it with Aspose.Cells, accesses the first worksheet, creates a Range object for D4:G4, unmerges that block, then iterates each cell to set its horizontal alignment to left before saving the workbook as output.xlsx. Exceptions are caught and reported.
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
                Console.WriteLine($"Input file \"{inputPath}\" not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the merged block covering D4:G4
            Aspose.Cells.Range mergedRange = sheet.Cells.CreateRange("D4:G4");
            mergedRange.UnMerge();

            // Restore individual cell alignment to left for each cell in the former merged range
            foreach (Cell cell in mergedRange)
            {
                // Retrieve the current style, modify alignment, and apply it back
                Style style = cell.GetStyle();
                style.HorizontalAlignment = TextAlignmentType.Left;
                cell.SetStyle(style);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
