// Title: Hide worksheets prefixed with "Temp" using Aspose.Cells for .NET
// Description: Loads an Excel file with Aspose.Cells, iterates all worksheets, hides those whose names start with "Temp" (case‑insensitive) by setting IsVisible = false, and saves the modified workbook.
// Keywords: Aspose.Cells hide worksheet | C# hide Temp sheets | iterate worksheets Aspose.Cells | set worksheet visibility | save workbook Aspose.Cells | case insensitive sheet name
// Common Searches: Aspose.Cells hide worksheet by name | C# hide all sheets starting with Temp | How to set worksheet visibility in Aspose.Cells | Save workbook after hiding sheets Aspose.Cells | Iterate through worksheets Aspose.Cells .NET
// Developer Intent: Hide every worksheet whose name begins with "Temp" and write the updated workbook to disk.
// Use Cases: Remove temporary analysis sheets before sharing a report. | Conceal intermediate calculation tabs in a financial model. | Prepare a clean template by hiding helper sheets that start with "Temp".
// AI Prompts: Generate C# code that uses Aspose.Cells to hide all worksheets whose names start with "Temp" and saves the workbook. | Show a LINQ‑based approach to filter and hide worksheets with a specific prefix in Aspose.Cells. | Explain how to unhide worksheets that were hidden by setting IsVisible = false in Aspose.Cells.

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, iterates all worksheets, hides those whose names start with "Temp" (case‑insensitive) by setting IsVisible = false, and saves the modified workbook.
class HideTempSheets
{
    static void Main()
    {
        // Load an existing workbook (create/load step)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // Hide worksheets whose names start with "Temp"
            if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
            {
                sheet.IsVisible = false; // hide the sheet
            }
        }

        // Save the modified workbook (save step)
        workbook.Save("output.xlsx");
    }
}
