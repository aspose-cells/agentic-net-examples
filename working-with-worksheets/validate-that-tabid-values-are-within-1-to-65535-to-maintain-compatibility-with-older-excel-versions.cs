// Title: Validate Worksheet TabId (1‑65535) with Aspose.Cells for .NET
// Description: Shows how to create a workbook, deliberately set an out‑of‑range TabId, loop through all worksheets, ensure each TabId is between 1 and 65,535, reset invalid entries, log the actions, and save the file.
// Keywords: Aspose.Cells TabId validation | worksheet TabId range .NET | Excel TabId limit 65535 | reset invalid TabId Aspose | TabId compatibility older Excel | C# Aspose.Cells workbook TabId
// Common Searches: Aspose.Cells check worksheet TabId range | how to fix TabId greater than 65535 in Excel using C# | validate TabId before saving workbook Aspose | set valid TabId for worksheets Aspose.Cells | Excel TabId out of range error solution
// Developer Intent: Ensure every worksheet's TabId stays within the 1‑65535 range and automatically correct any values that fall outside this limit.
// Use Cases: Automatically adjust TabId values that exceed Excel's maximum when generating reports programmatically. | Log worksheets with invalid TabId for compliance auditing before distribution. | Maintain compatibility with legacy Excel versions by enforcing valid TabId identifiers.
// AI Prompts: Create a C# method using Aspose.Cells that clamps each worksheet TabId to the 1‑65535 range and returns the names of sheets that were modified. | Generate code that validates TabId values, writes a correction log, and saves the workbook with a user‑specified filename. | Write a unit test in C# that verifies TabId handling for values below 1, within range, and above 65,535.

using System;
using Aspose.Cells;

namespace AsposeCellsTabIdValidation
{
    // Shows how to create a workbook, deliberately set an out‑of‑range TabId, loop through all worksheets, ensure each TabId is between 1 and 65,535, reset invalid entries, log the actions, and save the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example: set an invalid TabId (greater than 65535)
            sheet.TabId = 70000;

            // Validate TabId values for all worksheets
            foreach (Worksheet ws in workbook.Worksheets)
            {
                int tabId = ws.TabId;

                // TabId must be between 1 and 65535 inclusive
                if (tabId < 1 || tabId > 65535)
                {
                    // Adjust to a valid value (e.g., set to 1) and optionally log
                    Console.WriteLine($"Worksheet \"{ws.Name}\" has invalid TabId {tabId}. Resetting to 1.");
                    ws.TabId = 1;
                }
                else
                {
                    Console.WriteLine($"Worksheet \"{ws.Name}\" TabId {tabId} is valid.");
                }
            }

            // Save the workbook (lifecycle rule: save)
            string outputPath = "ValidatedTabIdWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
