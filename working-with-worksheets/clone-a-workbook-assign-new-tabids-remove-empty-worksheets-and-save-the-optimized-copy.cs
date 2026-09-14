// Title: Clone an Excel workbook, regenerate worksheet TabIds, purge empty sheets, and save the cleaned copy with Aspose.Cells for .NET
// AI Prompts: Write C# code that copies an existing Excel workbook, forces a new TabId for every worksheet, removes any worksheets that contain no data, and saves the cleaned workbook as an XLSX file using Aspose.Cells. | Generate a .NET method that clones a workbook, iterates through its worksheets to reset TabId values, deletes blank sheets, and writes the optimized workbook to disk with Aspose.Cells.
// Common Searches: Aspose.Cells how to copy a workbook and assign new TabId to each worksheet | C# remove blank worksheets after cloning an Excel file with Aspose.Cells | reset worksheet TabId programmatically in Aspose.Cells .NET | save optimized Excel workbook after deleting empty sheets using Aspose.Cells | clone Excel workbook and purge empty sheets in C#
// Tags: regenerate worksheet TabId after workbook copy | purge empty sheets during Excel workbook cloning | save cleaned workbook as XLSX using Aspose.Cells | C# Aspose.Cells workbook cloning with TabId reset | programmatic removal of blank worksheets .NET

using Aspose.Cells;
using System;
using System.IO;

// The program loads an existing Excel file, clones it into a new Workbook, forces a new TabId for each worksheet, removes any worksheets that are completely empty, and saves the resulting optimized workbook as optimized_copy.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            string originalPath = "original.xlsx";

            // Ensure the source file exists
            if (!File.Exists(originalPath))
            {
                Console.WriteLine($"Input file not found: {originalPath}");
                return;
            }

            // Load the original workbook
            Workbook originalWorkbook = new Workbook(originalPath);

            // Create an empty workbook to receive the copy
            Workbook clonedWorkbook = new Workbook();

            // Copy the original workbook into the new workbook
            originalWorkbook.Copy(clonedWorkbook);

            // Assign new TabIds to each worksheet (0 forces generation of a new unique TabId)
            foreach (Worksheet ws in clonedWorkbook.Worksheets)
            {
                ws.TabId = 0;
            }

            // Remove empty worksheets (iterate backwards for safe removal)
            for (int i = clonedWorkbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet ws = clonedWorkbook.Worksheets[i];
                bool isEmpty = ws.Cells.MaxDataRow < 0 && ws.Cells.MaxDataColumn < 0;
                if (isEmpty)
                {
                    clonedWorkbook.Worksheets.RemoveAt(i);
                }
            }

            // Save the optimized copy
            string optimizedPath = "optimized_copy.xlsx";
            clonedWorkbook.Save(optimizedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Optimized workbook saved to {optimizedPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception: {ex.Message}");
        }
    }
}
