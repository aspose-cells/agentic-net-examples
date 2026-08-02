// Title: Reorder Worksheets and Freeze Header Row in Aspose.Cells for .NET (C#)
// Description: Shows how to move worksheets to new positions with MoveTo and apply FreezePanes to the first row of every sheet in a workbook using Aspose.Cells for .NET, then saves the file as ReorderedAndFrozen.xlsx.
// Keywords: Aspose.Cells MoveTo | Aspose.Cells FreezePanes | C# reorder worksheets | freeze top row C# | Excel sheet ordering Aspose | header row freeze Aspose.Cells | workbook manipulation .NET
// Common Searches: Aspose.Cells reorder worksheets example | How to freeze first row in all sheets using Aspose.Cells | Move worksheet to first position C# Aspose | Apply FreezePanes to every worksheet Aspose.Cells | Change sheet order programmatically Aspose
// Developer Intent: The developer wants to programmatically change the sequence of worksheets and then lock the header row on each sheet so column titles stay visible during scrolling.
// Use Cases: Prepare a report where a summary sheet must appear first, then lock headers for easier data review. | Automate Excel export for large datasets, ensuring consistent sheet order and frozen header rows across all tabs. | Create printable workbooks where sheet sequence and frozen panes are required for a uniform layout.
// AI Prompts: Write C# code with Aspose.Cells that moves a worksheet named "Data" to the second position and freezes its top row. | Provide an example that iterates through all worksheets in a workbook and applies FreezePanes to keep the first row visible. | Explain how MoveTo and FreezePanes can be combined to reorder sheets and lock header rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to move worksheets to new positions with MoveTo and apply FreezePanes to the first row of every sheet in a workbook using Aspose.Cells for .NET, then saves the file as ReorderedAndFrozen.xlsx.
class ReorderAndFreezeDemo
{
    static void Main()
    {
        // Create a new workbook and add sample worksheets
        Workbook workbook = new Workbook();
        WorksheetCollection sheets = workbook.Worksheets;

        sheets.Add("First");
        sheets.Add("Second");
        sheets.Add("Third");
        sheets.Add("Fourth");

        // Reorder worksheets:
        // Move "Third" sheet to index 1 (second position)
        Worksheet thirdSheet = sheets["Third"];
        thirdSheet.MoveTo(1);

        // Move "Fourth" sheet to index 0 (first position)
        Worksheet fourthSheet = sheets["Fourth"];
        fourthSheet.MoveTo(0);

        // Freeze the top row of each worksheet to keep headers visible
        foreach (Worksheet ws in sheets)
        {
            // Freeze the first row (row 1) while leaving columns unfrozen
            // "A2" is the cell just below the header row
            ws.FreezePanes("A2", 1, 0);
        }

        // Save the modified workbook
        workbook.Save("ReorderedAndFrozen.xlsx");
    }
}
