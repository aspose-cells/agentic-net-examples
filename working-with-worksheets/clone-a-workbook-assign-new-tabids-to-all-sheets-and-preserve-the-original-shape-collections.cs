// Title: Clone an Excel workbook using Aspose.Cells .NET, assign new random TabIds to all worksheets, and preserve shape collections
// AI Prompts: Generate C# code that clones an existing .xlsx file with Aspose.Cells, assigns a random TabId to every worksheet in the clone, and saves the result while keeping all shapes intact. | Show how to copy a workbook via a MemoryStream in Aspose.Cells, modify the TabId property of each worksheet, and ensure shape objects are preserved.
// Common Searches: Aspose.Cells clone workbook and change worksheet TabId C# | preserve shapes when copying Excel file with Aspose.Cells .NET | assign unique TabId to worksheets after cloning workbook using Aspose.Cells | how to use MemoryStream to duplicate an Excel workbook in C# Aspose.Cells | retain shape objects while cloning workbook Aspose.Cells
// Tags: Aspose.Cells clone workbook via MemoryStream | assign random TabId to each worksheet Aspose.Cells | keep shape collections when duplicating Excel file | modify sheet TabId after workbook copy | C# generate unique TabId values for Excel sheets

using System;
using System.IO;
using Aspose.Cells;

// The example checks for the source file, loads it into a Workbook, clones it by saving to a MemoryStream and reloading, assigns a random integer TabId to each worksheet in the cloned workbook, and saves the cloned workbook as a new file while retaining the original shape collections.
class WorkbookCloneExample
{
    static void Main()
    {
        try
        {
            const string originalPath = "OriginalWorkbook.xlsx";
            const string clonedPath = "ClonedWorkbook.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Source file not found: {originalPath}");
                return;
            }

            // Load the original workbook
            Workbook originalWb = new Workbook(originalPath);

            // Clone the workbook by saving to a memory stream and reloading
            using (MemoryStream ms = new MemoryStream())
            {
                originalWb.Save(ms, SaveFormat.Xlsx);
                ms.Position = 0;
                Workbook clonedWb = new Workbook(ms);

                // Assign new TabIds (int) to all worksheets in the cloned workbook
                Random rnd = new Random();
                foreach (Worksheet sheet in clonedWb.Worksheets)
                {
                    sheet.TabId = rnd.Next(1, int.MaxValue);
                }

                // Save the cloned workbook
                clonedWb.Save(clonedPath);
                Console.WriteLine($"Cloned workbook saved to: {clonedPath}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
