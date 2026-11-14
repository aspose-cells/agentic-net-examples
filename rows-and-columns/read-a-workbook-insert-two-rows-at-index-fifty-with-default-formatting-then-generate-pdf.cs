// Title: Insert two rows at row 51 in an Excel worksheet and export the workbook to PDF using Aspose.Cells for .NET (C#)
// Common Searches: C# code to add two rows at a specific position in an Excel file with Aspose.Cells | How to insert rows with default formatting in Aspose.Cells before saving as PDF | Aspose.Cells example for inserting rows at index 50 and converting workbook to PDF | Convert modified Excel workbook to PDF after row insertion using Aspose.Cells in .NET | Insert rows at row 51 in Excel and export to PDF with Aspose.Cells C# tutorial
// Tags: insert rows Aspose.Cells C# | default row formatting Aspose.Cells | save workbook as PDF Aspose.Cells | worksheet row insertion before PDF export | Excel to PDF conversion after modifying rows C#

using System;
using Aspose.Cells;

// The sample loads an existing Excel file, inserts two rows at zero‑based index 50 (row 51) on the first worksheet using default formatting, and then saves the updated workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index if needed)
        Worksheet sheet = workbook.Worksheets[0];

        // Insert two rows at index 50 (zero‑based). This adds rows with default formatting.
        sheet.Cells.InsertRows(50, 2);

        // Save the modified workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
