// Title: Delete blank worksheets from an Excel workbook before converting to PDF with Aspose.Cells for .NET
// AI Prompts: Generate C# code that scans a Workbook, removes any worksheet that has no used cells, and then saves the remaining sheets as a PDF using Aspose.Cells. | Show how to use MaxDataRow and MaxDataColumn to detect empty sheets, delete them, and export the cleaned workbook to PDF in C#.
// Common Searches: aspnet remove empty worksheets before PDF export Aspose.Cells | c# detect and delete blank Excel sheets using MaxDataRow | skip blank worksheets when saving Excel as PDF with Aspose.Cells | avoid extra blank pages in PDF generated from Excel workbook | how to programmatically clean up workbook sheets prior to PDF conversion
// Tags: delete empty worksheets Aspose.Cells | export non‑empty sheets to PDF Aspose.Cells | MaxDataRow empty sheet detection C# | skip blank worksheets during PDF conversion | clean workbook before PDF export Aspose.Cells

using Aspose.Cells;
using System;
using System.Collections.Generic;

// Loads an Excel file, checks each worksheet’s MaxDataRow and MaxDataColumn to find sheets without data, removes those empty worksheets, and saves the remaining content as a PDF.
class Program
{
    static void Main()
    {
        // Load the workbook from an existing Excel file
        Workbook workbook = new Workbook("input.xlsx");

        // Identify indexes of worksheets that contain no data
        List<int> emptySheetIndexes = new List<int>();
        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet sheet = workbook.Worksheets[i];

            // MaxDataRow/MaxDataColumn are -1 when the sheet has no used cells
            if (sheet.Cells.MaxDataRow < 0 && sheet.Cells.MaxDataColumn < 0)
            {
                emptySheetIndexes.Add(i);
            }
        }

        // Remove empty worksheets starting from the highest index to avoid shifting
        for (int i = emptySheetIndexes.Count - 1; i >= 0; i--)
        {
            workbook.Worksheets.RemoveAt(emptySheetIndexes[i]);
        }

        // Save the workbook as PDF; only non‑empty worksheets will be rendered
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
