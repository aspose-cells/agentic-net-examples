// Title: Set 95% zoom on worksheets with over 500 rows and save the workbook as a PDF using Aspose.Cells for .NET
// AI Prompts: Loop through each worksheet, check Cells.MaxDataRow, set Worksheet.Zoom to 95 when rows exceed 500, then call Workbook.Save with SaveFormat.Pdf. | Determine the used row count per sheet, apply conditional zoom scaling, and generate a single PDF file from the workbook using C# and Aspose.Cells.
// Common Searches: Aspose.Cells C# set worksheet zoom based on row count before PDF export | How to apply 95% zoom to Excel sheets with more than 500 rows using Aspose.Cells | Export Excel workbook to PDF after adjusting zoom for large worksheets in .NET | Iterate all worksheets in Aspose.Cells and change zoom conditionally
// Tags: conditional worksheet zoom Aspose.Cells | apply 95% zoom to large sheets | export workbook as PDF Aspose.Cells | iterate worksheets by used row count | set Worksheet.Zoom C#

using System;
using Aspose.Cells;

// Loads an Excel workbook, loops through each worksheet, checks the number of used rows, sets the worksheet zoom to 95% when more than 500 rows are present, and saves the entire workbook as a single PDF file.
class Program
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Determine the number of rows that contain data
            // MaxDataRow is zero‑based, so add 1 to get the count
            int usedRowCount = sheet.Cells.MaxDataRow + 1;

            // Apply 95% zoom if the worksheet has more than 500 rows
            if (usedRowCount > 500)
            {
                sheet.Zoom = 95;
            }
        }

        // Export the entire workbook to a single PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
