// Title: Hide all worksheets whose names start with "Temp" and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Load an Excel workbook, loop through each Worksheet, set IsVisible = false when the worksheet name begins with "Temp" (case‑insensitive), then save the workbook to a new file. | Using Aspose.Cells for .NET, programmatically hide every sheet whose name matches a given prefix and write the updated workbook back to disk.
// Common Searches: Aspose.Cells C# hide worksheets that start with Temp | How to programmatically hide Excel sheets with a specific prefix using Aspose.Cells | Save workbook after changing worksheet visibility in C# Aspose.Cells | Iterate through worksheets and set IsVisible false based on name in Aspose.Cells | C# Aspose.Cells hide temporary sheets and save file
// Tags: hide worksheets by name prefix Aspose.Cells | set worksheet IsVisible property C# | iterate workbook worksheets Aspose.Cells | save modified Excel workbook Aspose.Cells | filter temporary sheets programmatically

using Aspose.Cells;
using System;

// The code loads 'input.xlsx', iterates over all worksheets, hides any sheet whose name starts with "Temp" (case‑insensitive) by setting IsVisible to false, and saves the updated workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Hide the worksheet if its name starts with "Temp"
            if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
            {
                sheet.IsVisible = false;
            }
        }

        // Save the modified workbook back to a file
        workbook.Save("output.xlsx");
    }
}
