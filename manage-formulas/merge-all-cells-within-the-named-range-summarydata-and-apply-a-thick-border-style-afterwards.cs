using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input workbook exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook that contains the named range "SummaryData"
            Workbook workbook = new Workbook(inputPath);
            Worksheet worksheet = workbook.Worksheets[0];

            // Obtain the range object for the named range
            Aspose.Cells.Range summaryRange = worksheet.Cells.CreateRange("SummaryData");

            // Merge all cells inside the named range
            summaryRange.Merge();

            // Convert the merged range to a UnionRange to apply border styling
            Aspose.Cells.Range baseRange = worksheet.Cells.CreateRange(summaryRange.RefersTo);
            UnionRange unionRange = baseRange.UnionRanges(new Aspose.Cells.Range[] { summaryRange });

            // Apply a thick black outline border around the merged range
            unionRange.SetOutlineBorders(CellBorderType.Thick, Color.Black);

            // Save the workbook with the changes
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}