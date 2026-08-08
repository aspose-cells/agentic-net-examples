// Title: Auto‑Fit All Columns After Merging Workbooks with Aspose.Cells for .NET (C#)
// Description: Loads two Excel files, merges them using Workbook.Combine, auto‑fits every column in each worksheet, and saves the resulting workbook.
// Keywords: Aspose.Cells combine workbooks | AutoFitColumns C# | merge Excel files .NET | adjust column width after combine | auto size columns Aspose | C# Excel workbook merge
// Common Searches: Aspose.Cells auto fit columns after combine | C# merge two Excel workbooks and auto size columns | How to auto‑fit all sheets after Workbook.Combine | AutoFitColumns for merged workbook Aspose | Combine workbooks and adjust column widths C#
// Developer Intent: Combine multiple Excel workbooks and automatically size columns for readability.
// Use Cases: Consolidate monthly reports into a single workbook with columns sized for clear presentation. | Create a unified financial statement from departmental files, ensuring consistent column widths after merging. | Prepare a merged dataset for analysis where each sheet’s columns are automatically optimized.
// AI Prompts: Generate C# code that merges three Excel workbooks with Aspose.Cells and applies AutoFitColumns to every worksheet before saving. | Explain how to auto‑fit a specific column range after using Workbook.Combine in Aspose.Cells. | Provide a step‑by‑step guide to merge workbooks, auto‑adjust column widths, and handle hidden sheets using Aspose.Cells.

using System;
using Aspose.Cells;

// Loads two Excel files, merges them using Workbook.Combine, auto‑fits every column in each worksheet, and saves the resulting workbook.
class Program
{
    static void Main()
    {
        // Load the first workbook (source)
        Workbook mergedWorkbook = new Workbook("FirstWorkbook.xlsx");

        // Load the second workbook to be combined
        Workbook secondWorkbook = new Workbook("SecondWorkbook.xlsx");

        // Merge the second workbook into the first one
        mergedWorkbook.Combine(secondWorkbook);

        // Auto‑fit all columns in every worksheet of the merged workbook
        foreach (Worksheet sheet in mergedWorkbook.Worksheets)
        {
            sheet.AutoFitColumns();
        }

        // Save the resulting workbook
        mergedWorkbook.Save("MergedWorkbook_AutoFit.xlsx");
    }
}
