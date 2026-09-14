// Title: Delete rows 10‑15 from an Excel worksheet and export the cleaned sheet to PDF using Aspose.Cells for .NET
// AI Prompts: Remove rows 10 through 15 from the first worksheet of an Excel file and save the cleaned sheet as a PDF with Aspose.Cells in C#. | Load an .xlsx workbook, delete a specific range of rows, and generate a PDF output using Aspose.Cells for .NET.
// Common Searches: how to delete specific rows in Excel using Aspose.Cells C# | export modified worksheet to PDF with Aspose.Cells .NET | remove rows 10-15 and save as PDF Aspose.Cells example
// Tags: Aspose.Cells delete rows range | Aspose.Cells export worksheet to PDF | C# remove specific rows Excel | Aspose.Cells PDF conversion after row deletion

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // for PdfSaveOptions if needed

// // This program loads "input.xlsx", deletes rows 10‑15 (zero‑based index 9, count 6) from the first worksheet, and saves the cleaned worksheet as "cleaned_output.pdf" using Aspose.Cells.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputFile = "cleaned_output.pdf";

        // Load the workbook from the existing file
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Delete rows 10 through 15.
        // Aspose.Cells uses zero‑based indexing, so row 10 is index 9.
        // Total rows to delete = 6 (10,11,12,13,14,15).
        worksheet.Cells.DeleteRows(9, 6);

        // Save the modified workbook as a PDF.
        // No special PDF options are required for this simple case.
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}
