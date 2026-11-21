// Title: Remove all slicers from every worksheet and save the workbook as PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to delete all slicers in an Excel file using Aspose.Cells | How to export an Excel workbook to PDF after removing slicers with Aspose.Cells | Iterate through worksheets and clear slicers programmatically in Aspose.Cells .NET | Aspose.Cells remove slicers from each sheet before PDF conversion | Sample for bulk slicer removal and PDF save in Aspose.Cells C#
// Tags: bulk slicer deletion Aspose.Cells | worksheet slicer removal C# | PDF conversion after slicer removal Aspose.Cells | slicer cleanup before PDF export | iterate worksheets remove slicers .NET

using Aspose.Cells;
using System;

// The example loads an existing Excel workbook, loops through each worksheet, removes every slicer by iterating the Slicers collection backwards, and then saves the modified workbook as a PDF file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Remove all slicers from the current worksheet
            // Iterate backwards to avoid index issues while removing
            for (int i = sheet.Slicers.Count - 1; i >= 0; i--)
            {
                sheet.Slicers.RemoveAt(i);
            }
        }

        // Export the modified workbook to PDF format
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
