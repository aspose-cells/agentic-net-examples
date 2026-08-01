// Title: Aspose.Cells C# – Delete Row 8 from an Excel worksheet and save as PDF
// Description: Loads an Excel file with Aspose.Cells, removes the eighth row (zero‑based index 7) from the first worksheet, and saves the updated workbook directly to PDF using SaveFormat.Pdf.
// Keywords: Aspose.Cells | C# | delete row | Excel row removal | row 8 | PDF conversion | SaveFormat.Pdf | worksheet manipulation | .NET | Excel to PDF | row deletion example
// Common Searches: Aspose.Cells delete row 8 C# | remove specific row from Excel using Aspose.Cells | convert Excel to PDF after deleting a row | C# delete Excel row and export to PDF | how to delete a row in Aspose.Cells
// Developer Intent: Remove the eighth row from the first worksheet of an Excel file and export the modified workbook as a PDF document.
// Use Cases: Strip a placeholder or header row before generating a client‑ready PDF report. | Automate cleanup of generated spreadsheets by deleting empty rows prior to distribution. | Prepare financial statements for regulatory submission by removing confidential rows and converting to PDF.
// AI Prompts: Write C# code with Aspose.Cells that deletes row 8 from the first worksheet and saves the workbook as a PDF. | Explain how to delete multiple consecutive rows in an Excel sheet using Aspose.Cells before converting to PDF. | Show how to add try‑catch error handling around row deletion and PDF saving with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Loads an Excel file with Aspose.Cells, removes the eighth row (zero‑based index 7) from the first worksheet, and saves the updated workbook directly to PDF using SaveFormat.Pdf.
class Program
{
    static void Main()
    {
        // Path to the source Excel file
        string sourcePath = "input.xlsx";

        // Path for the resulting PDF file
        string pdfPath = "output.pdf";

        // Load the workbook from the existing Excel file
        Workbook workbook = new Workbook(sourcePath);

        // Delete the 8th row (zero‑based index = 7) in the first worksheet
        workbook.Worksheets[0].Cells.DeleteRow(7);

        // Save the modified workbook as a PDF document
        workbook.Save(pdfPath, SaveFormat.Pdf);
    }
}
