// Title: C# example: Delete column Q from an Excel worksheet and export the result to PDF with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells code to remove a specific column from an Excel sheet | How to delete column Q in an .xlsx file using Aspose.Cells and then save as PDF | Aspose.Cells .NET delete column before converting worksheet to PDF | Remove Excel column programmatically in C# and generate PDF output | Aspose.Cells tutorial for column deletion and PDF export in .NET
// Tags: Aspose.Cells delete column C# | Excel column removal Aspose.Cells | convert worksheet to PDF Aspose.Cells | Aspose.Cells PDF export after column deletion | C# modify workbook then save as PDF

using System;
using Aspose.Cells;

// Loads input.xlsx, deletes column Q (zero‑based index 16) from the first worksheet using Aspose.Cells, and saves the modified workbook as output.pdf in PDF format.
class Program
{
    static void Main()
    {
        // Load the existing workbook from file
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet (adjust index or name as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete column Q (zero‑based index 16)
        worksheet.Cells.DeleteColumn(16);

        // Save the modified workbook as a PDF document
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
