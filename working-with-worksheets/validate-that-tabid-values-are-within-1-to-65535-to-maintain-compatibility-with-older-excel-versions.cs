// Title: Check and Correct Worksheet TabId Values (1‑65535) Using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an Excel workbook with Aspose.Cells, iterates every worksheet, verifies that each sheet's TabId is between 1 and 65535, logs any out‑of‑range IDs, and resets invalid values to a safe default before saving. | Create a reusable method in C# that validates worksheet TabId compatibility for older Excel versions, returns a list of sheets with invalid IDs, updates those TabId values, and persists the corrected workbook.
// Common Searches: how to validate worksheet tabid range with aspose.cells c# | reset invalid tabid values in excel workbook using asp.net | c# check tabid compatibility for Excel 97-2003 with Aspose.Cells | iterate worksheets and fix tabid out of range aspnet core | ensure excel file tabid stays within 1‑65535 using Aspose
// Tags: validate worksheet tabid range Aspose.Cells | reset out-of-range tabid c# | excel 97-2003 compatibility tabid | iterate worksheets tabid check Aspose | c# tabid validation Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an existing workbook, walks through each worksheet, checks that the TabId property falls within the 1‑65535 range required for legacy Excel compatibility, logs any violations, resets invalid TabId values to 1, and saves the corrected workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and validate their TabId values
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            int tabId = sheet.TabId;

            // TabId must be between 1 and 65535 for compatibility with older Excel versions
            if (tabId < 1 || tabId > 65535)
            {
                // Log the issue and reset to a safe default (e.g., 1)
                Console.WriteLine($"Worksheet '{sheet.Name}' has invalid TabId {tabId}. Resetting to 1.");
                sheet.TabId = 1;
            }
        }

        // Save the workbook after validation
        workbook.Save("output.xlsx");
    }
}
