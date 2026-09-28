// Title: How to clone an Excel workbook, assign new TabId values to every worksheet, and keep shape collections intact with Aspose.Cells for .NET
// AI Prompts: Generate C# code that uses Aspose.Cells to deep‑clone a workbook, set each worksheet's TabId to a sequential number starting at 1000, and save the result without losing any shape objects. | Write a program that loads an existing .xlsx file, calls Workbook.Copy to duplicate it, updates the TabId of all worksheets in the copy, and writes the cloned file while preserving all embedded shapes.
// Common Searches: how to copy a workbook and change worksheet tab ids with Aspose.Cells | preserve shapes when cloning an Excel file using Aspose.Cells .NET | assign new TabId to each worksheet after workbook copy | Aspose.Cells example for cloning workbook and updating TabId
// Tags: Workbook.Copy method for deep workbook duplication | modify Worksheet.TabId after workbook cloning | maintain shape objects while copying Excel workbook | sequential TabId assignment to worksheets .NET | duplicate Excel workbook with updated TabIds using Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// // Loads 'original.xlsx', creates a deep copy via Workbook.Copy, assigns new TabId values starting at 1000 to every worksheet in the cloned workbook, and saves it as 'cloned.xlsx' while preserving all existing shape collections.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "original.xlsx";
            const string destPath = "cloned.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException.
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the original workbook.
            Workbook originalWorkbook = new Workbook(sourcePath);

            // Create a new workbook and copy the contents of the original workbook.
            Workbook clonedWorkbook = new Workbook();
            clonedWorkbook.Copy(originalWorkbook);

            // Assign new, unique TabIds to every worksheet in the cloned workbook.
            int nextTabId = 1000;
            foreach (Worksheet sheet in clonedWorkbook.Worksheets)
            {
                sheet.TabId = nextTabId++;
            }

            // Save the cloned workbook.
            clonedWorkbook.Save(destPath);
            Console.WriteLine($"Cloned workbook saved to: {destPath}");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors.
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
