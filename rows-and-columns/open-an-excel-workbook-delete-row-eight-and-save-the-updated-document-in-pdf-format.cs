// Title: C# example: delete row 8 from the first worksheet of an Excel file and export to PDF using Aspose.Cells
// Common Searches: how to remove the 8th row from an Excel worksheet with Aspose.Cells in C# | Aspose.Cells C# code to delete a specific row and save as PDF | export modified Excel workbook to PDF after deleting a row using Aspose.Cells .NET | C# delete row 8 in Excel and convert to PDF with Aspose.Cells library | Aspose.Cells delete rows example and PDF output in .NET
// Tags: delete row 8 Aspose.Cells C# | remove specific worksheet row Aspose.Cells | export Excel to PDF Aspose.Cells .NET | save modified workbook as PDF C# Aspose.Cells | Aspose.Cells row deletion and PDF conversion | worksheet row removal and PDF export example

using System;
using Aspose.Cells;

// The sample loads an existing Excel workbook (input.xlsx) with Aspose.Cells, accesses the first worksheet, deletes the eighth row (zero‑based index 7), and then saves the updated workbook directly as a PDF file (output.pdf) using the SaveFormat.Pdf option.
class Program
{
    static void Main()
    {
        // Load the existing Excel workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (index 0)
        Worksheet sheet = workbook.Worksheets[0];

        // Delete row 8 (zero‑based index 7) – remove one row
        sheet.Cells.DeleteRows(7, 1);

        // Save the modified workbook as a PDF file
        workbook.Save("output.pdf", SaveFormat.Pdf);
    }
}
