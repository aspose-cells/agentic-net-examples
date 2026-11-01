// Title: Automatically assign the next TabId to a newly added worksheet in Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that scans all worksheets to find the highest TabId, adds a new sheet, and sets its TabId to the next integer. | Show how to programmatically assign a custom sequential TabId to a worksheet after adding it to a workbook with Aspose.Cells.
// Common Searches: aspnet how to get the highest TabId of worksheets in an Aspose.Cells workbook | c# add new worksheet and set TabId to max+1 using Aspose.Cells | aspose.cells sequential TabId assignment for newly created sheets | retrieve and increment worksheet TabId in a .xlsx file with Aspose.Cells
// Tags: set worksheet TabId Aspose.Cells .NET | retrieve maximum worksheet TabId Aspose.Cells | add worksheet with sequential TabId C# | incremental TabId assignment Aspose.Cells workbook | custom TabId for new worksheet Aspose.Cells

using Aspose.Cells;
using System;
using System.Linq;

// Loads an existing workbook, determines the highest TabId among its worksheets, adds a new worksheet, assigns the new sheet a TabId of max+1, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Determine the highest existing TabId among all worksheets
        int maxTabId = 0;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Assume Worksheet has a TabId property (int). If not, replace with appropriate property.
            int currentId = sheet.TabId;
            if (currentId > maxTabId)
                maxTabId = currentId;
        }

        // Add a new worksheet
        int newSheetIndex = workbook.Worksheets.Add();
        Worksheet newSheet = workbook.Worksheets[newSheetIndex];
        newSheet.Name = "NewSheet";

        // Assign an incremental TabId based on the highest existing identifier
        newSheet.TabId = maxTabId + 1;

        // Save the workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
