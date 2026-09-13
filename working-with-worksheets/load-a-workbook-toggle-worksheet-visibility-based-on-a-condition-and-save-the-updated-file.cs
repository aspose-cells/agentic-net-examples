// Title: Hide worksheets whose names start with a specific prefix and save the updated workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an Excel file with Aspose.Cells, hides every worksheet whose Name begins with "Temp", ensures all other sheets are visible, and saves the workbook to a new file. | Generate a C# program using Aspose.Cells that iterates through all worksheets, toggles the IsVisible property based on a custom condition (e.g., name prefix), and writes the modified workbook to disk.
// Common Searches: aspnet hide worksheets that start with Temp using Aspose.Cells | C# Aspose.Cells set worksheet visibility based on name condition | how to programmatically hide Excel sheets and save file with Aspose.Cells .NET | toggle worksheet IsVisible property for multiple sheets Aspose.Cells example
// Tags: hide worksheets by name prefix Aspose.Cells | set worksheet IsVisible property C# | conditional worksheet visibility Aspose.Cells | save modified workbook Aspose.Cells .NET

using Aspose.Cells;
using System;

// // Loads 'input.xlsx', hides any worksheet whose name starts with "Temp", ensures all other sheets are visible, and saves the result as 'output.xlsx' using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

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

        // Save the updated workbook to a new file
        workbook.Save("output.xlsx");
    }
}
