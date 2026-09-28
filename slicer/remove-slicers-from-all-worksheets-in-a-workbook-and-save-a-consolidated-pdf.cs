// Title: Remove slicers from every worksheet and export the workbook as a single PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that iterates over all worksheets, clears every slicer object, and then saves the workbook as one consolidated PDF file. | Show how to programmatically purge slicer collections from a workbook and perform a single‑PDF export using Aspose.Cells in .NET.
// Common Searches: C# Aspose.Cells how to remove slicer objects from every worksheet before PDF export | Export Excel workbook to one PDF without slicers using Aspose.Cells | Clear slicer collections in a .NET workbook programmatically | Aspose.Cells remove all slicers then save as consolidated PDF
// Tags: Aspose.Cells worksheet slicer purge | C# single PDF export after slicer cleanup | loop through worksheets to delete slicer objects | PDF generation without slicers Aspose.Cells | batch remove slicers from Excel workbook .NET

using System;
using Aspose.Cells;

// Loads an Excel file, iterates through each worksheet to clear its slicer collection, and saves the entire workbook as a single PDF document using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through all worksheets and remove all slicers
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // Clear the slicers collection for the current worksheet
            sheet.Slicers.Clear();
        }

        // Save the workbook as a consolidated PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
