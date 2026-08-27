// Title: How to set all slicers in an Excel workbook to non‑printable using Aspose.Cells for .NET (C#)
// AI Prompts: Generate C# code with Aspose.Cells that loads a workbook, iterates every worksheet, and sets each slicer's IsPrintable property to false before saving. | Provide a C# snippet that disables printing for all slicers in an existing Excel file using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# set slicer printable false for all worksheets | disable slicer printing in Excel file using Aspose.Cells .NET | iterate slicer collection and change IsPrintable property Aspose.Cells | batch update slicer printable flag in workbook Aspose.Cells C#
// Tags: Aspose.Cells slicer IsPrintable | C# disable slicer printing | Excel slicer non‑printable Aspose | batch update slicer properties .NET | iterate workbook slicer collection Aspose

using System;
using Aspose.Cells;
using Aspose.Cells.Slicers;

// Loads an Excel workbook, loops through each worksheet and its slicer collection, sets IsPrintable = false for every slicer, and saves the modified file.
class SetSlicersNonPrintable
{
    static void Main()
    {
        // Load an existing workbook (replace with your file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Access the slicer collection of the current worksheet
            SlicerCollection slicers = sheet.Slicers;

            // Iterate through each slicer and set its printable flag to false
            foreach (Slicer slicer in slicers)
            {
                slicer.IsPrintable = false;
            }
        }

        // Save the modified workbook (replace with your desired output path)
        workbook.Save("output.xlsx");
    }
}
