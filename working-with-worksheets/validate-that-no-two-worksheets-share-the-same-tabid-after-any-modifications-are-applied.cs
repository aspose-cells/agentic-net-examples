using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsTabIdValidation
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Add a few worksheets
                Worksheet sheet1 = workbook.Worksheets[0]; // default first sheet
                sheet1.Name = "Sheet1";
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Set TabId values (intentionally create a duplicate for demonstration)
                sheet1.TabId = 101;
                sheet2.TabId = 102;
                sheet3.TabId = 101; // duplicate TabId

                // Validate that all TabId values are unique
                try
                {
                    ValidateUniqueTabIds(workbook);
                }
                catch (InvalidOperationException ex)
                {
                    Console.WriteLine($"Validation error: {ex.Message}");
                    // Resolve duplicate by assigning a new unique TabId
                    ResolveDuplicateTabIds(workbook);
                    Console.WriteLine("Duplicate TabIds have been resolved.");
                }

                // Save the workbook if the target path is writable
                string outputPath = "ValidatedWorkbook.xlsx";
                try
                {
                    workbook.Save(outputPath);
                    Console.WriteLine($"Workbook saved to \"{outputPath}\".");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to save workbook: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Unexpected error: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks that no two worksheets in the given workbook share the same TabId.
        /// Throws InvalidOperationException if a duplicate is found.
        /// </summary>
        /// <param name="workbook">The workbook to validate.</param>
        static void ValidateUniqueTabIds(Workbook workbook)
        {
            HashSet<int> seenTabIds = new HashSet<int>();

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int tabId = sheet.TabId;

                // If the TabId has already been encountered, it's a duplicate
                if (!seenTabIds.Add(tabId))
                {
                    string message = $"Duplicate TabId detected: {tabId} on worksheet \"{sheet.Name}\".";
                    throw new InvalidOperationException(message);
                }
            }

            Console.WriteLine("All worksheet TabId values are unique.");
        }

        /// <summary>
        /// Resolves duplicate TabId values by assigning new unique identifiers.
        /// </summary>
        /// <param name="workbook">The workbook to fix.</param>
        static void ResolveDuplicateTabIds(Workbook workbook)
        {
            HashSet<int> usedIds = new HashSet<int>();
            int nextId = 1;

            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Find the next unused TabId
                while (usedIds.Contains(nextId))
                {
                    nextId++;
                }

                // If current TabId is already used, assign a new one
                if (!usedIds.Add(sheet.TabId))
                {
                    sheet.TabId = nextId;
                    usedIds.Add(nextId);
                    nextId++;
                }
                else
                {
                    // Current TabId is unique, keep it
                    // Ensure nextId starts above the highest used value
                    if (sheet.TabId >= nextId)
                    {
                        nextId = sheet.TabId + 1;
                    }
                }
            }
        }
    }
}