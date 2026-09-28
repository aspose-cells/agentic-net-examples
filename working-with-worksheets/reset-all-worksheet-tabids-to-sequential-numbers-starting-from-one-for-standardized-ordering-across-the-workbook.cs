// Title: Reset worksheet TabId values to sequential numbers with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an Excel file using Aspose.Cells, iterates through all worksheets, assigns TabId values starting at 1, and saves the modified workbook. | Create a .NET script with Aspose.Cells that renumbers the TabId property of every sheet in a workbook so the tabs are ordered 1,2,3,… and writes the result to a new file.
// Common Searches: how to renumber Excel sheet TabId using Aspose.Cells in C# | Aspose.Cells set TabId sequentially for all worksheets | C# code to reset worksheet tab identifiers to consecutive numbers | standardize tab order in an Excel workbook with Aspose.Cells
// Tags: Aspose.Cells sequential TabId assignment | C# worksheet TabId reset | Excel workbook reorder sheet tabs Aspose | TabId property usage in Aspose.Cells | programmatic worksheet ordering .NET

using System;
using Aspose.Cells;

// Loads an Excel workbook, iterates through each worksheet assigning TabId values 1, 2, 3…, and saves the updated file.
class Program
{
    static void Main()
    {
        // Path to the source workbook
        string inputPath = "input.xlsx";

        // Load the workbook
        Workbook workbook = new Workbook(inputPath);

        // Reset TabId for each worksheet to sequential numbers starting from 1
        int nextTabId = 1;
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.TabId = nextTabId;
            nextTabId++;
        }

        // Path for the updated workbook
        string outputPath = "output.xlsx";

        // Save the modified workbook
        workbook.Save(outputPath);
    }
}
