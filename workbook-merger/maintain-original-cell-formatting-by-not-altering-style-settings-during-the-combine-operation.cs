// Title: Combine Excel worksheets into a new workbook while preserving all cell styles with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an existing .xlsx file, creates a fresh Workbook, copies each worksheet using Worksheets.AddCopy so that formatting stays intact, and saves the merged workbook. | Demonstrate how to delete the default sheet from a newly created Aspose.Cells workbook before adding copied worksheets, ensuring original cell styles are not altered.
// Common Searches: Aspose.Cells C# copy worksheets without losing cell formatting | How to merge multiple sheets into one workbook while keeping styles in .NET | Preserve Excel cell styles when combining workbooks using Aspose.Cells AddCopy method
// Tags: Aspose.Cells Worksheets.AddCopy keep original formatting | C# merge Excel workbooks retain cell styles | combine worksheets into new workbook Aspose.Cells | remove default sheet from new Aspose.Cells workbook | copy Excel sheets without style changes .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads a source .xlsx file, creates a new Workbook, removes its default sheet, then iterates through the source worksheets and adds each one to the target using Worksheets.AddCopy, which copies values, formulas, and all formatting. Finally, the combined workbook is saved, ensuring that no cell style information is lost during the merge.
class CombineWorkbooksPreserveFormatting
{
    static void Main()
    {
        try
        {
            const string sourcePath = "SourceWorkbook.xlsx";
            const string outputPath = "CombinedWorkbook.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Create a new workbook that will receive the combined sheets (target)
            Workbook targetWorkbook = new Workbook(); // contains a default sheet

            // Load the source workbook containing sheets to be combined
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Remove the default sheet from the target if it is not needed
            if (targetWorkbook.Worksheets.Count == 1 && targetWorkbook.Worksheets[0].Name == "Sheet1")
                targetWorkbook.Worksheets.RemoveAt(0);

            // Iterate through each worksheet in the source workbook
            foreach (Worksheet srcSheet in sourceWorkbook.Worksheets)
            {
                // Add a copy of the source worksheet to the target workbook.
                // AddCopy(string sourceSheetName) copies all cell values, formulas, and formatting.
                targetWorkbook.Worksheets.AddCopy(srcSheet.Name);
            }

            // Save the combined workbook; original formatting of all cells is preserved.
            targetWorkbook.Save(outputPath);
            Console.WriteLine($"Workbooks combined successfully. Output saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
