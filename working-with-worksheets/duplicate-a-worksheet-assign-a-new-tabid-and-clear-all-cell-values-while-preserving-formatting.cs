// Title: Duplicate a worksheet, assign a new TabId, and clear cell values while preserving formatting using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that copies a worksheet, gives the copy a unique TabId, and clears only the cell contents while keeping all formatting with Aspose.Cells. | Write a .NET method to duplicate a specified sheet, set its TabId to the next highest value, and remove all data without affecting styles using Aspose.Cells.
// Common Searches: Aspose.Cells copy worksheet and set new TabId in C# | clear worksheet data but keep formatting Aspose.Cells .NET | how to assign a unique TabId to a duplicated sheet using Aspose.Cells
// Tags: duplicate worksheet with new TabId Aspose.Cells | clear worksheet contents preserving formatting .NET | assign unique TabId to copied sheet Aspose.Cells | copy sheet and reset values without losing styles

using Aspose.Cells;
using System;
using System.IO;

// The example loads an existing workbook, duplicates the first worksheet, assigns the copy a TabId higher than any existing sheet, clears all cell values while retaining formatting, and saves the result to a new file.
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

            // Index of the worksheet to duplicate (e.g., first worksheet)
            int sourceIndex = 0;
            Worksheet sourceSheet = workbook.Worksheets[sourceIndex];

            // Duplicate the worksheet
            int newIndex = workbook.Worksheets.AddCopy(sourceIndex);
            Worksheet duplicatedSheet = workbook.Worksheets[newIndex];

            // Assign a new unique TabId to the duplicated worksheet
            int maxTabId = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.TabId > maxTabId)
                    maxTabId = ws.TabId;
            }
            duplicatedSheet.TabId = maxTabId + 1;

            // Clear all cell values while preserving formatting
            // Use ClearContents overload that specifies the range to clear
            int maxRow = duplicatedSheet.Cells.MaxDataRow;
            int maxColumn = duplicatedSheet.Cells.MaxDataColumn;
            if (maxRow >= 0 && maxColumn >= 0)
            {
                duplicatedSheet.Cells.ClearContents(0, 0, maxRow + 1, maxColumn + 1);
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to \"{outputPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
