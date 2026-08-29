// Title: Auto‑fit all columns in each worksheet after merging two Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads two .xlsx files, merges them with Workbook.Combine, calls AutoFitColumns on every worksheet, and saves the result. | Show how to iterate through mergedWorkbook.Worksheets and apply AutoFitColumns to adjust column widths automatically. | Create a reusable method that accepts a list of Excel file paths, merges them into one workbook, auto‑fits columns across all sheets, and returns the saved file path.
// Common Searches: Aspose.Cells C# merge multiple Excel files and auto adjust column widths | auto fit columns in all sheets after combining workbooks with Aspose | C# example for Workbook.Combine followed by AutoFitColumns for each worksheet | how to automatically resize columns in a merged Excel workbook using Aspose.Cells | best practice for column width auto‑fit after workbook combine in .NET
// Tags: Workbook.Combine AutoFitColumns | auto‑fit columns post workbook combine | C# Aspose.Cells column width auto‑adjust | merge Excel workbooks and resize columns .NET | auto‑adjust column widths in combined workbook

using System;
using Aspose.Cells;

// The example loads source1.xlsx and source2.xlsx, merges them with Workbook.Combine, iterates through each worksheet to call AutoFitColumns, and saves the merged workbook as merged_autofit.xlsx.
class AutoFitMergedWorkbook
{
    static void Main()
    {
        // Load the first workbook which will serve as the base for merging
        Workbook mergedWorkbook = new Workbook("source1.xlsx");

        // Load the second workbook to be combined with the first one
        Workbook secondWorkbook = new Workbook("source2.xlsx");

        // Combine the second workbook into the first (mergedWorkbook)
        mergedWorkbook.Combine(secondWorkbook);

        // Apply AutoFit to all columns in every worksheet of the merged workbook
        foreach (Worksheet sheet in mergedWorkbook.Worksheets)
        {
            sheet.AutoFitColumns();
        }

        // Save the resulting workbook with auto‑fitted columns
        mergedWorkbook.Save("merged_autofit.xlsx", SaveFormat.Xlsx);
    }
}
