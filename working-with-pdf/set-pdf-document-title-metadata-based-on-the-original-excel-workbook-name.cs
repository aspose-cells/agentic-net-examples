using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class Program
{
    static void Main()
    {
        // Path to the original Excel workbook
        string excelPath = "source.xlsx";

        // Load the workbook (using the load rule)
        Workbook workbook = new Workbook(excelPath);

        // Set the FileName property so we can retrieve the original file name
        workbook.FileName = Path.GetFileName(excelPath);

        // Use the file name (without extension) as the PDF document title
        string title = Path.GetFileNameWithoutExtension(workbook.FileName);
        workbook.BuiltInDocumentProperties.Title = title;

        // Configure PDF save options to display the document title in the PDF window title bar
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            DisplayDocTitle = true
        };

        // Save the workbook as PDF (using the save rule)
        string pdfPath = "output.pdf";
        workbook.Save(pdfPath, pdfOptions);
    }
}

// Author note: This example loads an Excel file, copies its file name to the built‑in Title property,
// enables DisplayDocTitle, and saves the workbook as a PDF with the title reflected in the PDF metadata.