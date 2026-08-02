using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Switch the worksheet back to normal view (no page‑break outlines)
        sheet.IsPageBreakPreview = false;
        // Also ensure the view type is set to NormalView
        sheet.ViewType = ViewType.NormalView;

        // Save the workbook
        workbook.Save("NormalViewWorkbook.xlsx");
    }
}

// Author: Aspose.Cells .NET example code.