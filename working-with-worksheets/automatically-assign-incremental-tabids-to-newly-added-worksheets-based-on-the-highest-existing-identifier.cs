// Title: C# – Incremental Worksheet TabId Assignment with Aspose.Cells
// Description: Demonstrates a helper method that adds a new worksheet to an Aspose.Cells workbook, determines the highest existing TabId, and assigns the new sheet a TabId that is one greater. The sample sets an initial TabId on the default sheet, adds three additional sheets, prints each sheet's TabId, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells TabId | C# incremental TabId | add worksheet with custom TabId | Aspose.Cells workbook TabId example | auto‑assign TabId Aspose.Cells
// Common Searches: Aspose.Cells set worksheet TabId C# | increment TabId when adding sheets Aspose.Cells | how to get maximum TabId in a workbook Aspose.Cells | C# assign sequential TabId to new worksheets
// Developer Intent: Automatically give each newly added worksheet a unique, sequential TabId based on the current maximum TabId in the workbook.
// Use Cases: Maintain stable external references to worksheets by using predictable TabIds. | Create report workbooks where sheet identifiers must match database keys or other sequential IDs. | Automate generation of multiple sheets while ensuring no TabId collisions for downstream processing.
// AI Prompts: Generate a C# method that adds a worksheet to an Aspose.Cells workbook and sets its TabId to the next highest value. | Show how to iterate through all worksheets in a workbook to find the maximum TabId and then assign incremental TabIds to a list of new sheets. | Explain best practices for avoiding TabId conflicts when manually assigning TabIds in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdDemo
{
    // Demonstrates a helper method that adds a new worksheet to an Aspose.Cells workbook, determines the highest existing TabId, and assigns the new sheet a TabId that is one greater. The sample sets an initial TabId on the default sheet, adds three additional sheets, prints each sheet's TabId, and saves the workbook as an XLSX file.
    class Program
    {
        // Adds a worksheet with a TabId that is one greater than the current maximum TabId in the workbook.
        static Worksheet AddWorksheetWithIncrementalTabId(Workbook workbook, string sheetName)
        {
            // Add the worksheet using the provided Add(string) method.
            Worksheet newSheet = workbook.Worksheets.Add(sheetName);

            // Determine the highest TabId among existing worksheets.
            int maxTabId = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.TabId > maxTabId)
                {
                    maxTabId = ws.TabId;
                }
            }

            // Assign the new incremental TabId.
            newSheet.TabId = maxTabId + 1;

            return newSheet;
        }

        static void Main(string[] args)
        {
            // Create a new workbook (contains one default worksheet).
            Workbook workbook = new Workbook();

            // Optionally set a specific TabId for the default sheet to illustrate the logic.
            workbook.Worksheets[0].TabId = 100;

            // Add several new worksheets with automatically incremented TabIds.
            AddWorksheetWithIncrementalTabId(workbook, "SheetA");
            AddWorksheetWithIncrementalTabId(workbook, "SheetB");
            AddWorksheetWithIncrementalTabId(workbook, "SheetC");

            // Display TabId values for verification.
            foreach (Worksheet ws in workbook.Worksheets)
            {
                Console.WriteLine($"Worksheet \"{ws.Name}\" has TabId: {ws.TabId}");
            }

            // Save the workbook.
            workbook.Save("IncrementalTabIdDemo.xlsx");
        }
    }
}
