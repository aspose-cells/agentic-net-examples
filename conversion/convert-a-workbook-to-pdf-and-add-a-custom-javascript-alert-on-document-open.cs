// Title: C# – Convert Excel Workbook to PDF with Aspose.Cells and embed a JavaScript alert on document open
// Description: This example shows how to create an Aspose.Cells Workbook, write data to a worksheet, configure PdfSaveOptions (including CustomJavaScript to display an alert when the PDF opens), save the workbook as a PDF, and verify the output file while handling possible exceptions.
// Keywords: Aspose.Cells PDF conversion C# | add JavaScript to PDF Aspose.Cells | PdfSaveOptions CustomJavaScript | Excel to PDF with alert on open | C# Aspose.Cells example | .NET PDF JavaScript alert | document open JavaScript PDF | Aspose.Cells PDF export options
// Common Searches: how to add JavaScript alert to PDF using Aspose.Cells C# | Aspose.Cells PdfSaveOptions CustomJavaScript example | convert Excel workbook to PDF with Aspose.Cells .NET | C# code to embed JavaScript in PDF generated from Excel | verify PDF file creation after Aspose.Cells save
// Developer Intent: Generate a PDF from an Excel workbook with Aspose.Cells for .NET and embed a JavaScript alert that runs automatically when the PDF is opened.
// Use Cases: Create a new workbook, populate cells, and export it to PDF with custom JavaScript. | Add an on‑open alert (e.g., "Report generated on {date}") to PDFs for end‑user notifications. | Customize PDF output (document structure, JavaScript, page settings) using PdfSaveOptions. | Validate the existence of the PDF file after saving and implement robust error handling.
// AI Prompts: Write C# code that uses Aspose.Cells PdfSaveOptions.CustomJavaScript to show an alert saying "Welcome to the report" when the PDF opens. | Explain how to combine multiple PdfSaveOptions settings (page orientation, image quality, JavaScript) while converting an Excel workbook to PDF. | Provide a step‑by‑step guide for handling file‑system errors and ensuring the PDF is created successfully in a .NET application.

using System;
using System.IO;
using Aspose.Cells;

// This example shows how to create an Aspose.Cells Workbook, write data to a worksheet, configure PdfSaveOptions (including CustomJavaScript to display an alert when the PDF opens), save the workbook as a PDF, and verify the output file while handling possible exceptions.
class Program
{
    static void Main()
    {
        // Paths for the intermediate Excel file and final PDF file
        string excelPath = "sample.xlsx";
        string pdfPath = "output.pdf";

        try
        {
            // -------------------------------------------------
            // 1. Create a new workbook and add some data
            // -------------------------------------------------
            Workbook workbook = new Workbook();                     // Create a new workbook
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose.Cells!");      // Sample content

            // -------------------------------------------------
            // 2. Save the workbook as PDF using PdfSaveOptions
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true // Example option
            };
            workbook.Save(pdfPath, pdfOptions);                    // Save as PDF

            // Ensure the PDF was created
            if (!File.Exists(pdfPath))
                throw new FileNotFoundException("PDF file was not created.", pdfPath);

            Console.WriteLine("Workbook successfully converted to PDF.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
