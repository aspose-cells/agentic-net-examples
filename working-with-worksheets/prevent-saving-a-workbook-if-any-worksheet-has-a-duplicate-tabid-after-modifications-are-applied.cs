using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace PreventDuplicateTabIdSave
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Add a few worksheets for demonstration
            int sheet1Idx = workbook.Worksheets.Add();
            int sheet2Idx = workbook.Worksheets.Add();
            int sheet3Idx = workbook.Worksheets.Add();

            Worksheet sheet1 = workbook.Worksheets[sheet1Idx];
            Worksheet sheet2 = workbook.Worksheets[sheet2Idx];
            Worksheet sheet3 = workbook.Worksheets[sheet3Idx];

            // Assign TabId values (intentionally create a duplicate)
            sheet1.TabId = 101;
            sheet2.TabId = 102;
            sheet3.TabId = 101; // Duplicate TabId

            // Check for duplicate TabId values before saving
            if (HasDuplicateTabIds(workbook))
            {
                Console.WriteLine("Error: One or more worksheets have duplicate TabId values. Save operation aborted.");
                // Optionally, you could throw an exception instead of just returning.
                return;
            }

            // Save the workbook (lifecycle rule: save)
            string outputPath = "Workbook_NoDuplicateTabId.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }

        // Helper method to detect duplicate TabId values across all worksheets
        private static bool HasDuplicateTabIds(Workbook workbook)
        {
            HashSet<int> seenTabIds = new HashSet<int>();
            foreach (Worksheet ws in workbook.Worksheets)
            {
                // TabId defaults to 0 if not set; include it in the check as well
                int tabId = ws.TabId;
                if (seenTabIds.Contains(tabId))
                {
                    // Duplicate found
                    return true;
                }
                seenTabIds.Add(tabId);
            }
            // No duplicates
            return false;
        }
    }
}