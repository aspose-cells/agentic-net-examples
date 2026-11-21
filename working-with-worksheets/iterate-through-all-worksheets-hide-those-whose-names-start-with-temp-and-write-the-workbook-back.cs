// Title: Hide all worksheets whose names start with "Temp" using Aspose.Cells for .NET and save the workbook
// Common Searches: C# Aspose.Cells hide worksheets that start with Temp before saving | How to programmatically hide Excel sheets with a specific prefix using Aspose.Cells | Iterate through all worksheets in a workbook and set IsVisible false for Temp sheets in .NET | Save an Excel file after hiding temporary worksheets with Aspose.Cells for .NET
// Tags: hide worksheets by name prefix Aspose.Cells | set worksheet IsVisible property C# | iterate workbook worksheets Aspose.Cells | case-insensitive worksheet name filter Aspose.Cells | save modified Excel workbook Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an Excel workbook with Aspose.Cells, loops through each worksheet, hides any sheet whose name begins with "Temp" (case‑insensitive) by setting its IsVisible property to false, and then saves the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Check if the worksheet name starts with "Temp"
            if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
            {
                // Hide the worksheet
                sheet.IsVisible = false;
            }
        }

        // Save the modified workbook back to a file
        workbook.Save("output.xlsx");
    }
}
