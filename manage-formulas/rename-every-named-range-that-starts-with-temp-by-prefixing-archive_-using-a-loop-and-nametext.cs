// Title: How to rename Excel named ranges that start with "Temp" by adding an "Archive_" prefix using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells rename all named ranges beginning with Temp | Add Archive_ prefix to specific Excel named ranges programmatically | Loop through workbook defined names and change Name.Text in .NET | Batch rename temporary named ranges in an Excel file using Aspose.Cells | Programmatically update Excel named range prefixes with C#
// Tags: rename defined names with Aspose.Cells | add prefix to Excel defined names .NET | iterate workbook names Aspose.Cells | update Name.Text property for Excel ranges C# | batch rename temporary names Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an Excel workbook, iterates over all defined names, and for each name whose Text starts with "Temp" it prepends "Archive_" to the name before saving the workbook.
class Program
{
    static void Main()
    {
        // Load the workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Loop through all defined names in the workbook
        foreach (Name name in workbook.Worksheets.Names)
        {
            // If the name starts with "Temp", rename it by adding the "Archive_" prefix
            if (name.Text.StartsWith("Temp", StringComparison.Ordinal))
            {
                name.Text = "Archive_" + name.Text;
            }
        }

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
