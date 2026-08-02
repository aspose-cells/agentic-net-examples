// Title: C# Sample – Reset All Worksheet TabId Values to Sequential Numbers Using Aspose.Cells
// Description: This Aspose.Cells for .NET example loads an existing Excel workbook, iterates through its worksheets, assigns each sheet a TabId starting at 1 (making the tab order sequential), and saves the modified file. Adjusting TabId ensures the visible sheet order matches the intended sequence, useful after adding or removing sheets programmatically.
// Keywords: Aspose.Cells | C# | .NET | Worksheet TabId | reset TabId | sequential tab order | Excel workbook | programmatic sheet ordering | TabId property | sample code | GitHub example
// Common Searches: Aspose.Cells set worksheet TabId C# | reset Excel sheet TabId programmatically | sequential TabId Aspose.Cells example | change worksheet tab order .NET | how to renumber worksheet TabId with Aspose
// Developer Intent: Assign sequential TabId numbers to every worksheet in a workbook.
// Use Cases: Ensure the visual tab order matches logical sheet sequence before distributing a workbook. | Re‑order tabs after inserting or deleting sheets via automation. | Prepare workbooks for downstream processes that rely on consecutive TabId values. | Standardize tab ordering across multiple workbooks in a batch operation.
// AI Prompts: Write C# code that uses Aspose.Cells to set each worksheet's TabId to a consecutive number starting from 1 and save the workbook. | Explain the effect of the TabId property on Excel sheet ordering and how to modify it with Aspose.Cells. | Add comprehensive error handling to the TabId reset example, including file‑not‑found and permission checks. | Create a PowerShell script that calls a .NET assembly to renumber worksheet TabIds using Aspose.Cells.

using System;
using Aspose.Cells;

// This Aspose.Cells for .NET example loads an existing Excel workbook, iterates through its worksheets, assigns each sheet a TabId starting at 1 (making the tab order sequential), and saves the modified file. Adjusting TabId ensures the visible sheet order matches the intended sequence, useful after adding or removing sheets programmatically.
class ResetWorksheetTabIds
{
    static void Main()
    {
        // Load the workbook you want to modify
        string inputPath = "input.xlsx";
        Workbook workbook = new Workbook(inputPath);

        // Assign sequential TabId values starting from 1
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];
            sheet.TabId = i + 1;
        }

        // Save the workbook with updated TabIds
        string outputPath = "output.xlsx";
        workbook.Save(outputPath);
    }
}
