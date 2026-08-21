// Title: C# – Auto‑Increment Worksheet TabId in Aspose.Cells for .NET
// Description: Demonstrates how to add worksheets to an Aspose.Cells workbook and automatically assign each new sheet a sequential TabId by locating the current maximum TabId and setting the new value to max + 1. The workbook is saved as IncrementalTabIdDemo.xlsx.
// Keywords: Aspose.Cells | .NET | C# | Worksheet TabId | auto increment TabId | sequential TabId | add worksheet programmatically | unique TabId example | workbook sample code | GitHub Aspose.Cells example
// Common Searches: Aspose.Cells assign incremental TabId C# | auto increment worksheet TabId .NET | how to set unique TabId for new sheets in Aspose.Cells | C# code to get max TabId across worksheets | sample project for TabId sequencing Aspose.Cells
// Developer Intent: Add new worksheets and ensure each receives a unique, sequential TabId based on the highest existing identifier.
// Use Cases: Maintain predictable tab order when generating workbooks dynamically. | Synchronize worksheet IDs with external databases or APIs that expect sequential numbers. | Create user‑friendly Excel files where TabId reflects the creation sequence for easier navigation.
// AI Prompts: Generate a reusable method GetNextTabId(Workbook wb) that returns the next sequential TabId. | Refactor the loop to use LINQ for finding the maximum TabId in a workbook. | Explain how to detect and resolve duplicate TabId conflicts when inserting worksheets in parallel.

using System;
using Aspose.Cells;

// Demonstrates how to add worksheets to an Aspose.Cells workbook and automatically assign each new sheet a sequential TabId by locating the current maximum TabId and setting the new value to max + 1. The workbook is saved as IncrementalTabIdDemo.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Initialize the first worksheet's TabId
        Worksheet firstSheet = workbook.Worksheets[0];
        firstSheet.TabId = 1;

        // Add new worksheets and assign incremental TabIds
        for (int i = 0; i < 3; i++)
        {
            // Add a new worksheet with a unique name
            Worksheet newSheet = workbook.Worksheets.Add($"Sheet_{i + 2}");

            // Find the current maximum TabId among all worksheets
            int maxTabId = 0;
            foreach (Worksheet ws in workbook.Worksheets)
            {
                if (ws.TabId > maxTabId)
                    maxTabId = ws.TabId;
            }

            // Set the new worksheet's TabId to max + 1
            newSheet.TabId = maxTabId + 1;
        }

        // Save the workbook
        workbook.Save("IncrementalTabIdDemo.xlsx");
    }
}
