// Title: How to automatically assign the next TabId to a newly added worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that scans all worksheets in a Workbook, determines the highest TabId, adds a new worksheet, and sets its TabId to the next integer using Aspose.Cells. | Create a reusable C# method that accepts a Workbook object, adds a worksheet with an incremented TabId based on existing sheets, and returns the newly created worksheet.
// Common Searches: Aspose.Cells C# find maximum worksheet TabId and assign next value | auto increment TabId when adding new sheet with Aspose.Cells .NET | C# code to set sequential TabId for a newly added worksheet in Aspose.Cells | retrieve highest TabId from workbook and add sheet with next TabId Aspose.Cells
// Tags: increment TabId Aspose.Cells | max worksheet TabId retrieval .NET | add worksheet sequential TabId | Aspose.Cells workbook TabId management

using Aspose.Cells;
using System;

// The example loads an existing workbook, iterates through its worksheets to locate the highest TabId, adds a new worksheet, assigns the new sheet a TabId that is one greater than the maximum, optionally renames the sheet, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load the existing workbook (replace with your actual file path)
        var workbook = new Workbook("input.xlsx");

        // Find the highest TabId among all existing worksheets
        int maxTabId = 0;
        foreach (Worksheet ws in workbook.Worksheets)
        {
            // Assuming Worksheet has a TabId property (int)
            if (ws.TabId > maxTabId)
                maxTabId = ws.TabId;
        }

        // Add a new worksheet to the workbook
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];

        // Assign an incremental TabId based on the highest existing identifier
        newSheet.TabId = maxTabId + 1;

        // Optionally give the new sheet a name
        newSheet.Name = $"Sheet{newSheetIndex + 1}";

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
