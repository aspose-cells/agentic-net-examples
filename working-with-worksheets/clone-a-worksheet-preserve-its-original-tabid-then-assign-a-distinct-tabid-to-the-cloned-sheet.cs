// Title: Aspose.Cells .NET – Clone a Worksheet, Preserve Original TabId, Set New TabId
// Description: Demonstrates how to create a workbook, assign a specific TabId to the first worksheet, clone the sheet with Worksheets.AddCopy, retain the original TabId value, give the cloned sheet a different TabId, and save the result as an XLSX file using C#.
// Keywords: Aspose.Cells clone worksheet | Worksheet TabId property | preserve TabId on copy | set custom TabId Aspose.Cells | AddCopy TabId .NET | C# Aspose.Cells worksheet duplication | unique TabId for cloned sheet
// Common Searches: clone worksheet keep TabId Aspose.Cells | change TabId of copied worksheet C# | Aspose.Cells AddCopy preserve TabId | how to assign new TabId after worksheet clone | C# example TabId property Aspose.Cells
// Developer Intent: Copy an existing worksheet, retain its original TabId for reference, and assign a distinct TabId to the duplicate sheet.
// Use Cases: Generate monthly report tabs by cloning a template sheet and incrementing TabId for each month. | Create versioned data sheets where the base TabId stays constant while each version receives a unique TabId for tracking. | Duplicate configuration worksheets for different environments, preserving the master TabId and assigning environment‑specific TabIds to the copies.
// AI Prompts: Write C# code with Aspose.Cells that clones a worksheet and sets a custom TabId on the clone. | Explain the behavior of the TabId property when using Worksheets.AddCopy in Aspose.Cells for .NET. | Provide a sample that clones a worksheet, keeps the original TabId, and assigns a new TabId based on an offset calculation.

using System;
using Aspose.Cells;

// Demonstrates how to create a workbook, assign a specific TabId to the first worksheet, clone the sheet with Worksheets.AddCopy, retain the original TabId value, give the cloned sheet a different TabId, and save the result as an XLSX file using C#.
class WorksheetCloneTabIdDemo
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the original worksheet (first sheet) and set a known TabId
        Worksheet originalSheet = workbook.Worksheets[0];
        originalSheet.Name = "OriginalSheet";
        originalSheet.TabId = 100; // example original TabId

        // Clone the original worksheet using AddCopy (preserves data and formatting)
        int clonedIndex = workbook.Worksheets.AddCopy(originalSheet.Name);
        Worksheet clonedSheet = workbook.Worksheets[clonedIndex];
        clonedSheet.Name = "ClonedSheet";

        // Preserve the original TabId value
        int originalTabId = originalSheet.TabId;

        // Assign a distinct TabId to the cloned sheet (ensure it differs from the original)
        clonedSheet.TabId = originalTabId + 1;

        // Save the workbook with the cloned sheet
        workbook.Save("ClonedWorksheetTabIdDemo.xlsx");
    }
}
