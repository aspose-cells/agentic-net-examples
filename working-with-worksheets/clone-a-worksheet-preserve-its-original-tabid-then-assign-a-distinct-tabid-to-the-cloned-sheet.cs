// Title: How to clone an Excel worksheet and assign a unique TabId using Aspose.Cells for .NET
// Common Searches: Aspose.Cells .NET clone worksheet and change TabId | C# duplicate Excel sheet with new TabId using Aspose.Cells | preserve original TabId while copying worksheet in Aspose.Cells | assign different TabId to copied worksheet Aspose.Cells C# | AddCopy worksheet TabId handling Aspose.Cells .NET
// Tags: Aspose.Cells worksheet duplication API | C# set distinct TabId on duplicated worksheet | handle TabId conflict after worksheet copy | prevent duplicate sheet name in Aspose.Cells | retrieve worksheet index after AddCopy

using System;
using System.IO;
using Aspose.Cells;

// Loads a workbook, stores the original worksheet's TabId, clones the sheet with AddCopy, assigns a new distinct TabId to the cloned sheet, optionally renames it to avoid name clashes, and saves the updated workbook.
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
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (or any desired worksheet)
            Worksheet originalSheet = workbook.Worksheets[0];

            // Preserve the original TabId
            int originalTabId = originalSheet.TabId;

            // Clone the worksheet; AddCopy returns the index of the new sheet
            int clonedSheetIndex = workbook.Worksheets.AddCopy(originalSheet.Index);

            // Retrieve the cloned worksheet using its index
            Worksheet clonedSheet = workbook.Worksheets[clonedSheetIndex];

            // Assign a distinct TabId to the cloned sheet
            clonedSheet.TabId = originalTabId + 1;

            // Optionally rename the cloned sheet to avoid name conflicts
            clonedSheet.Name = originalSheet.Name + "_Clone";

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
