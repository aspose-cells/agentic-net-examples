// Title: Apply Page Break Preview view to all worksheets in an Excel workbook using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# set page break preview for every worksheet in a workbook | How to change worksheet view to PageBreakPreview before saving with Aspose.Cells .NET | Batch update view type to PageBreakPreview for all sheets using Aspose.Cells API | Prepare Excel file for printing by enabling Page Break Preview on all worksheets in C# | Set ViewType.PageBreakPreview for each worksheet in Aspose.Cells
// Tags: Aspose.Cells set worksheet viewtype | page break preview view Aspose.Cells | apply viewtype to all worksheets .NET | prepare Excel workbook for printing Aspose | batch worksheet view update Aspose.Cells

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook, iterates through every worksheet, sets each sheet's ViewType to PageBreakPreview, and saves the file, enabling a print‑ready layout across the entire workbook.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Set each worksheet to Page Break Preview view
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ViewType = ViewType.PageBreakPreview;
        }

        // Save the workbook with the updated view settings
        workbook.Save("output.xlsx");
    }
}
