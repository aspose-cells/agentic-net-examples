// Title: Set Page Break Preview view for every worksheet in an Excel workbook with Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file using Aspose.Cells, changes each worksheet's ViewType to PageBreakPreview, and saves the updated file. | Create a reusable function that accepts a Workbook object, iterates through its Worksheets collection, sets ViewType to PageBreakPreview, and returns the modified workbook.
// Common Searches: asp.net aspose.cells enable page break preview on all worksheets | c# programmatically switch Excel sheets to PageBreakPreview before printing | how to apply page break preview to every sheet in a workbook using Aspose.Cells | Aspose.Cells .NET set view type to PageBreakPreview for multiple worksheets | set worksheet view mode to page break preview with C# Aspose.Cells API
// Tags: Aspose.Cells set worksheet view type | PageBreakPreview view for Excel worksheets | C# change Excel sheet view mode | prepare workbook for printing with page break preview | apply viewtype pagebreakpreview using Aspose.Cells

using System;
using Aspose.Cells;

// // Loads an Excel file, iterates through each worksheet to set its ViewType to PageBreakPreview, and saves the updated workbook.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Apply Page Break Preview to every worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Set the view type to PageBreakPreview
            sheet.ViewType = ViewType.PageBreakPreview;
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
