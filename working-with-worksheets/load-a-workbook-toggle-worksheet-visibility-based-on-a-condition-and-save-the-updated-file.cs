// Title: Hide worksheets whose names start with a prefix and save the workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# hide worksheets that start with 'Temp' prefix | How to toggle worksheet visibility based on name pattern in Aspose.Cells | Saving workbook after changing worksheet IsVisible property with Aspose.Cells | Iterate through worksheets collection to hide specific sheets using Aspose.Cells .NET | Conditional worksheet visibility example in C# Aspose.Cells
// Tags: conditional worksheet visibility Aspose.Cells | hide worksheets by name prefix C# | set worksheet IsVisible property Aspose.Cells | persist workbook after visibility update Aspose.Cells | iterate worksheets collection .NET

using Aspose.Cells;
using System;

// Loads input.xlsx, hides any worksheet whose name begins with "Temp" by setting IsVisible to false, ensures other sheets are visible, and saves the updated file as output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Toggle worksheet visibility based on a condition.
        // Example condition: hide worksheets whose name starts with "Temp"
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name.StartsWith("Temp"))
            {
                // Hide the worksheet
                sheet.IsVisible = false;
            }
            else
            {
                // Ensure the worksheet is visible
                sheet.IsVisible = true;
            }
        }

        // Save the updated workbook
        workbook.Save("output.xlsx");
    }
}
