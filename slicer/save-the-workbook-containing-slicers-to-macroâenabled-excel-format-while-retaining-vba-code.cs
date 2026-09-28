// Title: Save an Excel workbook that contains slicers to a macro‑enabled .xlsm file while keeping VBA code using Aspose.Cells for .NET
// AI Prompts: Load an .xlsx file with slicers and write it out as an .xlsm, ensuring the VBA project is retained, using Aspose.Cells in C#. | Demonstrate how to export a slicer‑enabled workbook to macro‑enabled format without losing any VBA macros with Aspose.Cells for .NET. | Show the C# code to convert a workbook that includes slicers into a macro‑enabled Excel file while preserving all embedded VBA.
// Common Searches: Aspose.Cells C# save workbook with slicers as .xlsm preserving VBA | how to keep VBA macros when converting Excel file with slicers to macro enabled using Aspose | convert xlsx containing slicers to xlsm without losing VBA code in .NET | saving slicer‑enabled Excel workbook as macro‑enabled file with Aspose.Cells
// Tags: Aspose.Cells save slicer workbook as xlsm | retain VBA macros when exporting to macro‑enabled Excel | export workbook with slicers to .xlsm using C# | macro‑enabled Excel file creation Aspose.Cells

using System;
using Aspose.Cells;

// // Loads a workbook that contains slicers and saves it as a macro‑enabled .xlsm file, preserving any embedded VBA code using Aspose.Cells for .NET.
class SaveWorkbookWithSlicers
{
    static void Main()
    {
        // Load the workbook that contains slicers (any supported format)
        Workbook workbook = new Workbook("input.xlsx");

        // Save the workbook as a macro‑enabled file, preserving any VBA code
        workbook.Save("output.xlsm", SaveFormat.Xlsm);
    }
}
