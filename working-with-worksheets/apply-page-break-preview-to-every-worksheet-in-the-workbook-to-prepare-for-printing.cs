using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (replace with Load if needed)
        Workbook workbook = new Workbook();

        // Enable Page Break Preview for every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.IsPageBreakPreview = true; // Show sheet in Page Break Preview mode
        }

        // Save the workbook to a file
        workbook.Save("PageBreakPreview_Output.xlsx");
    }
}

// Author: Aspose.Cells .NET example – applies page break preview to all worksheets.