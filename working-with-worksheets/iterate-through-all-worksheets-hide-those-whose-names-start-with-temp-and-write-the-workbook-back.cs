// Title: Hide worksheets starting with “Temp” in a .NET workbook using Aspose.Cells and save it
// Description: Loads an existing workbook, loops through all worksheets, hides any sheet whose name begins with "Temp" (case‑insensitive) by setting IsVisible to false, and saves the updated file.
// Keywords: Aspose.Cells hide worksheet | C# hide worksheet by name | iterate worksheets Aspose.Cells | worksheet visibility Aspose.Cells | case‑insensitive worksheet name | save workbook Aspose.Cells | temporary sheets Excel | Aspose.Cells .NET
// Common Searches: Aspose.Cells hide worksheet that starts with Temp | C# hide Excel sheets by prefix using Aspose.Cells | How to set worksheet IsVisible false Aspose.Cells | Save workbook after hiding sheets Aspose.Cells | Iterate all worksheets Aspose.Cells C#
// Developer Intent: Programmatically hide every worksheet whose name begins with “Temp” and persist the changes to a new Excel file.
// Use Cases: Remove internal calculation tabs before sharing a report with clients. | Automatically conceal temporary analysis sheets generated during batch processing. | Prepare a clean workbook for distribution while keeping helper sheets hidden. | Reduce visual clutter in exported Excel files by hiding auxiliary worksheets.
// AI Prompts: Write C# code that uses Aspose.Cells to hide all worksheets whose names start with "Temp" and then saves the workbook. | Explain how to perform a case‑insensitive name check when hiding sheets with Aspose.Cells. | Show how to unhide worksheets that were hidden because their names start with "Temp" using Aspose.Cells. | Generate a script that processes multiple Excel files, hides temporary sheets in each, and saves the results.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an existing workbook, loops through all worksheets, hides any sheet whose name begins with "Temp" (case‑insensitive) by setting IsVisible to false, and saves the updated file.
    class HideTempSheets
    {
        static void Main()
        {
            // Load an existing workbook (create rule)
            Workbook workbook = new Workbook("input.xlsx");

            // Iterate through all worksheets
            for (int i = 0; i < workbook.Worksheets.Count; i++)
            {
                Worksheet sheet = workbook.Worksheets[i];

                // Hide the worksheet if its name starts with "Temp"
                if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
                {
                    // Set visibility to false (hide)
                    sheet.IsVisible = false;
                }
            }

            // Save the modified workbook (save rule)
            workbook.Save("output.xlsx");
        }
    }
}
