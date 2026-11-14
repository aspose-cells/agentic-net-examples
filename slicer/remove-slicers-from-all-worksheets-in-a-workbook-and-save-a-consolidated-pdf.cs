// Title: How to remove all slicers from every worksheet and export the workbook as a single PDF using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# delete slicers from all worksheets before PDF conversion | remove slicer objects from Excel workbook programmatically with Aspose.Cells | export Excel to one PDF file after clearing slicers using .NET | C# code to clear worksheet slicers and generate consolidated PDF with Aspose.Cells | how to batch remove slicers in a workbook and save as PDF in Aspose.Cells
// Tags: Aspose.Cells remove slicers from worksheets | clear slicer collections C# Aspose.Cells | save workbook as single PDF Aspose.Cells | batch slicer deletion in Excel .NET | consolidated PDF export after slicer removal

using System;
using Aspose.Cells;

// The example loads an Excel workbook, iterates through each worksheet to clear any slicers, and then saves the modified workbook as a single consolidated PDF using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all slicers from each worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // If the worksheet contains slicers, clear them
            if (sheet.Slicers.Count > 0)
            {
                sheet.Slicers.Clear();
            }
        }

        // Save the modified workbook as a single consolidated PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
