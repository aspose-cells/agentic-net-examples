// Title: Hide Excel worksheets whose names start with "Temp" and save the workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code that loads an Excel file with Aspose.Cells, iterates over each worksheet, hides any sheet whose Name begins with "Temp", and saves the modified workbook to a new file. | Write a method in C# using Aspose.Cells for .NET that sets the IsVisible property to false for all worksheets whose names have the prefix "Temp" and then writes the workbook out. | Create a C# example demonstrating how to programmatically hide temporary worksheets (named with "Temp") and export the updated workbook with Aspose.Cells.
// Common Searches: C# Aspose.Cells hide worksheets that start with Temp and save file | how to set worksheet visibility based on name prefix using Aspose.Cells .NET | iterate through workbook worksheets and hide Temp sheets in C# | Aspose.Cells hide temporary Excel sheets before exporting | filter out worksheets with Temp* prefix using Aspose.Cells API
// Tags: hide worksheets by name prefix Aspose.Cells C# | set worksheet IsVisible property Aspose.Cells | iterate workbook worksheets collection Aspose.Cells | save modified Excel workbook Aspose.Cells .NET | temporary sheet visibility Aspose.Cells

using System;
using Aspose.Cells;

// The program loads input.xlsx with Aspose.Cells, iterates all worksheets, hides any worksheet whose name starts with "Temp" (case‑insensitive) by setting IsVisible to false, and saves the result as output.xlsx.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Hide worksheets whose names start with "Temp"
            if (sheet.Name.StartsWith("Temp", StringComparison.OrdinalIgnoreCase))
            {
                sheet.IsVisible = false;
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
