// Title: Clone an Excel worksheet and assign a different TabId using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# clone worksheet and set a new TabId | preserve original TabId when copying a sheet with Aspose.Cells | assign unique TabId to duplicated worksheet in Aspose.Cells .NET | avoid TabId conflict after worksheet copy using Aspose.Cells | AddCopy method example with custom TabId in C#
// Tags: worksheet cloning Aspose.Cells | custom TabId assignment Excel sheet | AddCopy method Aspose.Cells C# | prevent TabId duplication Aspose.Cells | unique worksheet TabId generation .NET

using System;
using System.IO;
using Aspose.Cells;

// The example loads or creates an Excel workbook, sets a known TabId on the original worksheet, clones the sheet using Worksheets.AddCopy, determines an unused TabId, assigns it to the cloned worksheet, and saves the workbook.
class Program
{
    static void Main()
    {
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        Workbook workbook = null;

        // Load existing workbook or create a new one if the file does not exist.
        if (File.Exists(inputPath))
        {
            try
            {
                workbook = new Workbook(inputPath);
                Console.WriteLine($"Loaded workbook from '{inputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
                return;
            }
        }
        else
        {
            // Create a simple workbook with a single worksheet.
            workbook = new Workbook();
            Worksheet ws = workbook.Worksheets[0];
            ws.Name = "Original";
            ws.Cells["A1"].PutValue("Sample data");
            // Assign a known TabId for demonstration.
            ws.TabId = 100;
            Console.WriteLine($"Input file not found. Created new workbook with sheet '{ws.Name}' (TabId = {ws.TabId}).");
        }

        if (workbook == null)
        {
            Console.WriteLine("Workbook could not be initialized.");
            return;
        }

        // Reference the original worksheet (first sheet).
        Worksheet original = workbook.Worksheets[0];
        int originalTabId = original.TabId;
        Console.WriteLine($"Original worksheet '{original.Name}' TabId = {originalTabId}");

        // Clone the original worksheet.
        Worksheet cloned;
        try
        {
            // AddCopy returns the index of the new sheet in older API versions.
            int clonedIndex = workbook.Worksheets.AddCopy(original.Name);
            cloned = workbook.Worksheets[clonedIndex];
            cloned.Name = original.Name + "_Clone";
            Console.WriteLine($"Cloned worksheet created with name '{cloned.Name}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error cloning worksheet: {ex.Message}");
            return;
        }

        // Assign a distinct TabId to the clone, ensuring no conflict.
        int newTabId = originalTabId + 1;
        while (IsTabIdInUse(workbook, newTabId))
        {
            newTabId++;
        }
        cloned.TabId = newTabId;
        Console.WriteLine($"Cloned worksheet '{cloned.Name}' assigned new TabId = {cloned.TabId}");

        // Save the modified workbook.
        try
        {
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }

    // Helper method to verify whether a TabId is already used in the workbook.
    private static bool IsTabIdInUse(Workbook wb, int tabId)
    {
        foreach (Worksheet ws in wb.Worksheets)
        {
            if (ws.TabId == tabId)
                return true;
        }
        return false;
    }
}
