// Title: Delete column F from an Excel worksheet and export the workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# delete column F from Excel and save as PDF | remove specific column from worksheet using Aspose.Cells .NET | how to export modified Excel workbook to PDF after column deletion in C# | C# code to delete a column by index in Aspose.Cells and generate PDF | Aspose.Cells example for deleting column and converting to PDF
// Tags: delete column Aspose.Cells C# | export Excel to PDF Aspose.Cells | remove worksheet column by index | save modified workbook as PDF .NET | Aspose.Cells column deletion example | convert Excel workbook to PDF C#

using System;
using Aspose.Cells;

// This example loads an existing Excel file, deletes column F (the sixth column) from the first worksheet using Aspose.Cells, and then saves the workbook directly as a PDF document.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Delete column F (zero‑based index 5) from the first worksheet
        Worksheet sheet = workbook.Worksheets[0];
        sheet.Cells.DeleteColumn(5);

        // Save the modified workbook as PDF
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
