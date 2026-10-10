// Title: How to embed a PDF/A‑1a report as an OLE object in an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads a PDF/A‑1a file and inserts it as an OLE object at a specified cell range in an existing .xlsx workbook. | Show how to add error handling for missing workbook or PDF files and set custom dimensions and a name for the embedded PDF OLE object using Aspose.Cells. | Demonstrate saving the workbook after embedding the PDF/A‑1a report and verifying that the OLE object appears correctly in Excel.
// Common Searches: aspocells c# embed pdf a1a as ole object in existing workbook | how to attach a pdf file to an excel sheet using aspose.cells .net | set position and size of embedded pdf ole object in excel with c# | error handling for missing pdf or xlsx when embedding with aspose.cells | embed pdf/a-1a compliance report into excel using aspose.cells api
// Tags: aspose.cells embed pdf oleobject | c# embed pdf/a-1a into xlsx | oleobject placement row column aspose.cells | excel workbook pdf attachment aspnet | aspose.cells error handling missing files

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;   // Required for OleObject

// The example loads an existing Excel workbook, reads a PDF/A‑1a file into a byte array, adds it as an embedded OLE object at a defined row and column with custom dimensions and a name, handles missing file errors, and saves the workbook with the PDF attached.
class Program
{
    static void Main()
    {
        try
        {
            // Paths for the main workbook and the PDF to embed
            string mainWorkbookPath = "MainDocument.xlsx";
            string pdfAttachmentPath = "Report_PDF_A1a.pdf";

            // Verify that required files exist
            if (!File.Exists(mainWorkbookPath))
                throw new FileNotFoundException($"Main workbook not found: {mainWorkbookPath}");
            if (!File.Exists(pdfAttachmentPath))
                throw new FileNotFoundException($"Attachment PDF not found: {pdfAttachmentPath}");

            // Load the existing workbook
            Workbook workbook = new Workbook(mainWorkbookPath);
            Worksheet sheet = workbook.Worksheets[0]; // use the first worksheet

            // Read the PDF file bytes
            byte[] pdfData = File.ReadAllBytes(pdfAttachmentPath);

            // Define placement for the embedded OLE object (row, column, height, width)
            int row = 5;          // zero‑based row index
            int column = 2;       // zero‑based column index
            int height = 100;     // in pixels
            int width = 100;      // in pixels

            // Add an embedded OLE object representing the PDF using the byte array
            // Add returns the index of the newly added object
            int oleIndex = sheet.OleObjects.Add(row, column, height, width, pdfData);
            OleObject ole = sheet.OleObjects[oleIndex];
            ole.Name = "PDF_A1a_Report";   // optional name for the OLE object

            // Save the workbook with the embedded PDF
            string outputPath = "MainDocument_WithAttachment.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
