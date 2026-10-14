// Title: Hide worksheets containing a specific keyword and make others visible with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to open an Excel workbook, hide every worksheet whose name includes a given substring, set all remaining sheets to visible, and save the result to a new file. | Adapt the sample to accept the keyword and output path as variables, toggle worksheet visibility accordingly, and write the updated workbook using Aspose.Cells.
// Common Searches: Aspose.Cells C# hide worksheets whose name contains a specific keyword | Programmatically change Excel sheet visibility based on sheet name using Aspose.Cells | How to set IsVisible property for all worksheets in a workbook with Aspose.Cells | Save a modified Excel file to a different location after toggling sheet visibility in C#
// Tags: Aspose.Cells hide worksheets by name | C# set worksheet IsVisible property | conditional worksheet visibility Aspose.Cells | save workbook to new file Aspose.Cells | iterate over worksheets Aspose.Cells C#

using Aspose.Cells;
using System;

// Loads 'input.xlsx', hides any worksheet whose name contains the keyword "Hidden", ensures all other worksheets are visible, and saves the modified workbook as 'output.xlsx' using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet and toggle its visibility
        // Example condition: hide worksheets whose name contains "Hidden"
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            if (sheet.Name.Contains("Hidden"))
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

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
