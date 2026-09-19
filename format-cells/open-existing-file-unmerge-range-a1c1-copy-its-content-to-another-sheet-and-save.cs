// Title: Unmerge cells A1:C1 and copy the range to a new worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Load an existing workbook, unmerge the range A1:C1 on the first sheet, create a new worksheet named "CopySheet", copy the unmerged cells (values, formulas, formatting) to A1:C1 of the new sheet, and save the file. | Using Aspose.Cells for C#, demonstrate how to programmatically break a merged header row, duplicate its content to a newly added sheet, and persist the changes as a new Excel file.
// Common Searches: Aspose.Cells C# how to break merged cells and move them to another sheet | Copy header row after unmerging with Aspose.Cells .NET | Programmatic way to duplicate A1:C1 range to new worksheet in Excel using Aspose | Add worksheet and transfer cell values and styles with Aspose.Cells C#
// Tags: unmerge merged cells A1:C1 Aspose.Cells | copy cell range to another worksheet Aspose.Cells | create additional worksheet and transfer formatting .NET | read and write Excel file Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// // Loads "input.xlsx", unmerges the merged range A1:C1 on the first worksheet, adds a new worksheet called "CopySheet", copies the unmerged cells (including values, formulas, and formatting) to A1:C1 of the new sheet, and saves the result as "output.xlsx".
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

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Source worksheet (first sheet)
            Worksheet sourceSheet = workbook.Worksheets[0];

            // Unmerge the range A1:C1 using a Range object
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:C1");
            sourceRange.UnMerge();

            // Add a new worksheet to receive the copied content
            int newSheetIndex = workbook.Worksheets.Add();
            Worksheet destSheet = workbook.Worksheets[newSheetIndex];
            destSheet.Name = "CopySheet";

            // Define the destination range
            AsposeRange destRange = destSheet.Cells.CreateRange("A1:C1");

            // Copy values, formulas, formatting, etc. from source to destination
            destRange.Copy(sourceRange);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
