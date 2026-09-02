// Title: How to Validate and Auto‑Correct Worksheet TabId Values (1‑65535) with Aspose.Cells for .NET
// AI Prompts: Generate C# code using Aspose.Cells that loops through every worksheet in an Excel file, checks the TabId, and clamps any value outside 1‑65535 to the nearest valid limit. | Write a reusable method in C# that receives a Workbook object and throws an InvalidOperationException if any worksheet’s TabId is less than 1 or greater than 65535, leveraging Aspose.Cells APIs. | Create a console application that loads a workbook, validates each sheet’s TabId against the 1‑65535 range, optionally corrects invalid IDs, and saves the file, using Aspose.Cells for .NET.
// Common Searches: aspocells c# check worksheet tabid range 1-65535 | how to clamp invalid Excel sheet TabId with Aspose.Cells | C# program to enforce TabId limits for compatibility with older Excel versions | throw exception for out‑of‑range TabId in Aspose.Cells workbook | validate Excel worksheet TabId values using Aspose.Cells .NET library
// Tags: Aspose.Cells TabId range enforcement | C# adjust worksheet TabId | correct out‑of‑range TabId Aspose.Cells | Excel sheet TabId compatibility 1‑65535 | check TabId values .NET Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel workbook, iterates through each worksheet, verifies that the TabId is within the 1‑65535 range, optionally clamps invalid values or throws an exception, and saves the corrected workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and validate TabId values
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;

            // Check if TabId is outside the allowed range (1 to 65535)
            if (tabId < 1 || tabId > 65535)
            {
                // Option 1: Adjust to the nearest valid value
                sheet.TabId = Math.Max(1, Math.Min(tabId, 65535));

                // Option 2: Throw an exception to signal invalid data
                // throw new InvalidOperationException(
                //     $"Worksheet \"{sheet.Name}\" has an invalid TabId ({tabId}). Valid range is 1‑65535.");
            }
        }

        // Save the workbook after validation (or correction)
        workbook.Save("output.xlsx");
    }
}
