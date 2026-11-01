// Title: Unmerge cells T5:U5 and recalculate dependent formulas in an existing Excel workbook using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells unmerge a merged range and refresh dependent formulas | Unmerge cells T5:U5 in an existing .xlsx and recalculate totals with Aspose.Cells | Aspose.Cells .NET load workbook, unmerge specific cells, recalculate all formulas | How to programmatically unmerge a range and update formulas in Excel using Aspose.Cells | Save modified workbook after unmerging cells with Aspose.Cells C#
// Tags: unmerge merged cells Aspose.Cells C# | recalculate workbook formulas Aspose.Cells | load existing workbook and save as new file Aspose.Cells | modify merged cells Aspose.Cells | Excel formula refresh after unmerge Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing Excel file, unmerges the merged range T5:U5, triggers a full formula recalculation to update dependent totals, and saves the result to a new workbook, handling errors gracefully.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Unmerge the cells in the range T5:U5
            Aspose.Cells.Range range = sheet.Cells.CreateRange("T5:U5");
            range.UnMerge();

            // Recalculate all formulas so dependent totals are updated
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            workbook.Save(outputPath);

            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
