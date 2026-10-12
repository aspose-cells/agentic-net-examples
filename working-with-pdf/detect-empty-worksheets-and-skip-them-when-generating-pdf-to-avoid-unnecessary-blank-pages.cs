// Title: Remove blank worksheets from an Excel file before converting to PDF using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that scans each worksheet, filters out those that have no data, and then exports the remaining sheets to a PDF file. | Create a method in C# that identifies worksheets without any content in an Aspose.Cells workbook and excludes them from the PDF conversion process.
// Common Searches: how to skip blank sheets when converting Excel to PDF with Aspose.Cells | remove empty worksheets before saving workbook as PDF using Aspose.Cells for .NET | Aspose.Cells detect empty worksheet and exclude from PDF output | C# Aspose.Cells delete worksheets with no data | skip sheets without data during Aspose.Cells PDF export
// Tags: Aspose.Cells workbook remove empty sheets | Aspose.Cells PDF conversion ignore blank worksheets | C# identify blank worksheet Aspose.Cells | Aspose.Cells clean up workbook before PDF export | Aspose.Cells skip sheets without data

using Aspose.Cells;
using System;
using System.Collections.Generic;

// The program loads an Excel workbook, detects worksheets that contain no data, removes those blank sheets, and then saves the cleaned workbook as a PDF using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the Excel workbook (load rule)
        Workbook workbook = new Workbook("input.xlsx");

        // Identify empty worksheets
        List<int> emptySheetIndexes = new List<int>();
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // If there are no data rows/columns, the sheet is empty
            if (sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0)
            {
                emptySheetIndexes.Add(i);
                continue;
            }

            // Verify that all cells are actually empty
            bool hasData = false;
            for (int row = 0; row <= sheet.Cells.MaxDataRow && !hasData; row++)
            {
                for (int col = 0; col <= sheet.Cells.MaxDataColumn; col++)
                {
                    if (!string.IsNullOrEmpty(sheet.Cells[row, col].StringValue))
                    {
                        hasData = true;
                        break;
                    }
                }
            }

            if (!hasData)
                emptySheetIndexes.Add(i);
        }

        // Remove empty worksheets starting from the highest index
        for (int i = emptySheetIndexes.Count - 1; i >= 0; i--)
        {
            workbook.Worksheets.RemoveAt(emptySheetIndexes[i]);
        }

        // Save the workbook as PDF (save rule)
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
