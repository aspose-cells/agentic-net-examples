// Title: C# – Detect duplicate worksheet TabId after copying workbooks with Aspose.Cells
// Description: Loads a source workbook, copies all its worksheets to a new workbook using Workbook.Worksheets.AddCopy, then scans the destination workbook for repeated TabId values and reports any duplicates before saving.
// Keywords: Aspose.Cells | C# | TabId | duplicate worksheet ID | AddCopy | copy worksheets | Workbook | detect duplicate sheet identifiers | Excel automation | sheet ID conflict
// Common Searches: Aspose.Cells detect duplicate TabId after copying sheets | C# check for repeated worksheet IDs in destination workbook | AddCopy duplicate TabId issue | how to find duplicate sheet identifiers with Aspose.Cells | verify unique TabId values after workbook copy
// Developer Intent: Identify any worksheet TabId that occurs more than once after copying sheets from one workbook to another.
// Use Cases: Validate uniqueness of TabId values before performing formula updates or data merges. | Log duplicate TabId occurrences to troubleshoot reference errors in Excel automation. | Prevent runtime exceptions in applications that rely on a one‑to‑one mapping between worksheets and TabId.
// AI Prompts: Write C# code using Aspose.Cells to copy all worksheets from a source workbook to a destination workbook and list duplicate TabId values with their sheet names. | Create a reusable method that returns a collection of duplicate TabId entries after copying worksheets with Aspose.Cells. | Explain how Aspose.Cells assigns TabId to worksheets and why AddCopy might produce duplicate IDs.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace SheetIdDuplicationCheck
{
    // Loads a source workbook, copies all its worksheets to a new workbook using Workbook.Worksheets.AddCopy, then scans the destination workbook for repeated TabId values and reports any duplicates before saving.
    class Program
    {
        static void Main()
        {
            // Load the source workbook (existing file)
            Workbook sourceWorkbook = new Workbook("Source.xlsx");   // uses the provided constructor

            // Load (or create) the destination workbook
            Workbook destWorkbook = new Workbook();                 // empty workbook

            // Copy all worksheets from source to destination using AddCopy (by index)
            for (int i = 0; i < sourceWorkbook.Worksheets.Count; i++)
            {
                // AddCopy creates a new sheet in destWorkbook that is a copy of the source sheet at index i
                destWorkbook.Worksheets.AddCopy(i);
            }

            // After copying, check for duplicate TabId values in the destination workbook
            var tabIdMap = new Dictionary<int, List<string>>(); // TabId -> list of sheet names

            foreach (Worksheet sheet in destWorkbook.Worksheets)
            {
                int tabId = sheet.TabId; // internal sheet identifier
                if (!tabIdMap.ContainsKey(tabId))
                {
                    tabIdMap[tabId] = new List<string>();
                }
                tabIdMap[tabId].Add(sheet.Name);
            }

            // Report any TabId that appears more than once (potential duplication)
            bool duplicatesFound = false;
            foreach (var kvp in tabIdMap)
            {
                if (kvp.Value.Count > 1)
                {
                    duplicatesFound = true;
                    Console.WriteLine($"Duplicate TabId {kvp.Key} found in sheets: {string.Join(", ", kvp.Value)}");
                }
            }

            if (!duplicatesFound)
            {
                Console.WriteLine("No duplicate TabId values detected.");
            }

            // Save the destination workbook (uses the provided Save method)
            destWorkbook.Save("DestinationWithCopies.xlsx");
        }
    }
}
