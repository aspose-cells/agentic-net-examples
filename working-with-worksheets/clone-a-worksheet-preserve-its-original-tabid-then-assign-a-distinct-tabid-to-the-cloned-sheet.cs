// Title: Clone an Excel worksheet with Aspose.Cells for .NET while preserving its original TabId and assigning a unique new TabId
// AI Prompts: Create a C# function that uses Aspose.Cells to copy a worksheet, retains the original TabId, and assigns a non‑conflicting TabId to the new sheet. | Develop a reusable utility that clones any worksheet in a .NET workbook and automatically finds the next available TabId.
// Common Searches: Aspose.Cells C# clone worksheet preserving TabId | set new TabId for copied Excel sheet with Aspose | avoid TabId collisions when duplicating worksheets in .NET | how to generate a free TabId for a cloned worksheet using Aspose.Cells
// Tags: Aspose.Cells AddCopy worksheet cloning | preserve original TabId on worksheet copy | generate unique TabId for cloned sheet | C# workbook duplicate sheet TabId handling | Aspose.Cells TabId conflict resolution

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file, clones the first worksheet using AddCopy, keeps the source TabId, computes a non‑conflicting TabId for the clone, renames the cloned sheet, and saves the workbook.
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

            // Load the workbook from the input file
            Workbook workbook = new Workbook(inputPath);

            // Select the worksheet to clone (e.g., the first worksheet)
            Worksheet originalSheet = workbook.Worksheets[0];

            // Preserve the original TabId
            int originalTabId = originalSheet.TabId;

            // Clone the worksheet using AddCopy (returns the index of the new sheet)
            int clonedIndex = workbook.Worksheets.AddCopy(originalSheet.Index);
            Worksheet clonedSheet = workbook.Worksheets[clonedIndex];
            clonedSheet.Name = "Cloned_" + originalSheet.Name;

            // Assign a distinct TabId to the cloned sheet, ensuring no conflict
            int newTabId = originalTabId + 1;
            while (TabIdExists(workbook, newTabId))
            {
                newTabId++;
            }
            clonedSheet.TabId = newTabId;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle unexpected errors gracefully
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }

    // Helper method to check if a TabId already exists in the workbook
    static bool TabIdExists(Workbook wb, int tabId)
    {
        foreach (Worksheet ws in wb.Worksheets)
        {
            if (ws.TabId == tabId)
                return true;
        }
        return false;
    }
}
