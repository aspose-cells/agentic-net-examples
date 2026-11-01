// Title: Apply Page Break Preview view to every worksheet in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, sets ViewType to PageBreakPreview for each worksheet, and saves the workbook. | Show how to loop through Workbook.Worksheets and change the view mode to PageBreakPreview before exporting the file. | Demonstrate enabling page break preview on all sheets in a workbook programmatically with Aspose.Cells for .NET.
// Common Searches: how to enable page break preview for all sheets using Aspose.Cells C# | Aspose.Cells set view type to PageBreakPreview for every worksheet in a workbook | C# code to apply page break preview mode to an entire Excel file before saving | batch change worksheet view to PageBreakPreview with Aspose.Cells .NET
// Tags: Aspose.Cells set worksheet view to PageBreakPreview | C# loop through workbook worksheets change view type | enable page break preview before workbook save | batch update worksheet view mode Aspose.Cells | change Excel sheet view to PageBreakPreview .NET

using Aspose.Cells;
using System;

// // Loads an Excel workbook, iterates over each worksheet, sets its ViewType to PageBreakPreview, and saves the updated file.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply Page Break Preview to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ViewType = ViewType.PageBreakPreview;
        }

        // Save the modified workbook (replace with desired output path)
        workbook.Save("output.xlsx");
    }
}
