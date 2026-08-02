// Title: C# – Validate Unique Worksheet TabId Values in Aspose.Cells
// Description: Creates a workbook, adds three sheets, intentionally sets duplicate TabId values, then scans all worksheets with a HashSet to report any repeated TabId and confirms uniqueness before saving the file.
// Keywords: Aspose.Cells TabId validation | C# duplicate worksheet TabId | unique worksheet TabId Aspose | .NET Excel TabId conflict | detect repeated TabId cells
// Common Searches: Aspose.Cells check duplicate TabId C# | how to ensure unique worksheet TabId in .NET | C# code to find repeated TabId in Excel workbook | validate worksheet TabId uniqueness before save | detect TabId conflicts Aspose.Cells
// Developer Intent: Verify that no two worksheets share the same TabId after any changes.
// Use Cases: Run the validator after adding or renaming sheets to catch TabId collisions before exporting. | Integrate the check into an automated Excel report generator to prevent UI tab errors. | Extend the sample to automatically reassign new TabId values when duplicates are found.
// AI Prompts: Generate C# code that automatically assigns new unique TabId values to worksheets with duplicates in an Aspose.Cells workbook. | Show how to log duplicate TabId detections to a file instead of the console using Aspose.Cells. | Create an NUnit test that confirms ValidateUniqueTabIds correctly identifies and reports duplicate TabId entries.

using System;
using System.Collections.Generic;
using Aspose.Cells;

namespace AsposeCellsTabIdValidation
{
    // Creates a workbook, adds three sheets, intentionally sets duplicate TabId values, then scans all worksheets with a HashSet to report any repeated TabId and confirms uniqueness before saving the file.
    public class TabIdValidator
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the default first worksheet
                Worksheet sheet1 = workbook.Worksheets[0];
                sheet1.Name = "Sheet1";

                // Add a second worksheet
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

                // Add a third worksheet
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // Intentionally set duplicate TabId values for demonstration
                sheet1.TabId = 100;
                sheet2.TabId = 200;
                sheet3.TabId = 100; // Duplicate of sheet1

                // Validate that all worksheets have unique TabId values
                ValidateUniqueTabIds(workbook);

                // Save the workbook (adjust the path as needed)
                workbook.Save("TabIdValidationResult.xlsx");
                Console.WriteLine("Workbook saved successfully.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred during validation: {ex.Message}");
            }
        }

        private static void ValidateUniqueTabIds(Workbook workbook)
        {
            // Use a HashSet to track encountered TabId values
            HashSet<int> seenTabIds = new HashSet<int>();
            bool duplicateFound = false;

            // Iterate through all worksheets in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                int currentTabId = sheet.TabId;

                // Check if the TabId has already been seen
                if (!seenTabIds.Add(currentTabId))
                {
                    // Duplicate detected
                    duplicateFound = true;
                    Console.WriteLine($"Duplicate TabId detected: Worksheet \"{sheet.Name}\" has TabId {currentTabId} which is already used.");
                }
            }

            if (!duplicateFound)
            {
                Console.WriteLine("All worksheets have unique TabId values.");
            }
            else
            {
                // Optionally, resolve duplicates by assigning new unique TabIds
                // Here we simply report the issue; resolution logic can be added as needed.
                Console.WriteLine("Validation completed: duplicates exist.");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            TabIdValidator.Run();
        }
    }
}
