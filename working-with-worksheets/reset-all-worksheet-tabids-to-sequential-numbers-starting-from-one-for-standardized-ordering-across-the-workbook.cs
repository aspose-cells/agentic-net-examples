// Title: Renumber all worksheet TabId values sequentially starting at 1 using Aspose.Cells for .NET
// AI Prompts: Loop through each worksheet in a workbook, assign its TabId property an incrementing integer beginning with 1, and save the file with Aspose.Cells in C#. | Update the TabId of every sheet in an Excel workbook to create a continuous sequence (1,2,3…) using Aspose.Cells for .NET.
// Common Searches: C# Aspose.Cells how to set worksheet TabId sequentially | renumber Excel sheet tabs after deleting worksheets with Aspose.Cells | programmatically reset worksheet tab order using Aspose.Cells .NET | assign incremental TabId to all worksheets in a workbook Aspose.Cells | standardize Excel worksheet TabId values with C# Aspose.Cells
// Tags: set worksheet TabId Aspose.Cells | sequential TabId assignment .NET | renumber Excel sheet tabs C# | standardize worksheet tab order Aspose.Cells | update worksheet TabId property programmatically

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates over each worksheet to assign TabId values 1, 2, 3…, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Assign sequential TabIds starting from 1
        int tabId = 1;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.TabId = tabId;
            tabId++;
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
