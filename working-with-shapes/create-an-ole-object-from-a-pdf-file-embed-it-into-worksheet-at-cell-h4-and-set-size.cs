// Title: C# – Embed a PDF as an OLE object in cell H4 using Aspose.Cells
// Description: Demonstrates how to create a workbook, read a PDF file, add an OLE object placeholder at cell H4, set its width and height to 200 px, embed the PDF data, assign the PDF ProgID, and save the file with Aspose.Cells for .NET.
// Keywords: Aspose.Cells embed PDF | C# OLE object Excel | add OLE object cell H4 | set OLE object size Aspose | ProgID PDF Aspose.Cells | embed PDF in worksheet .NET
// Common Searches: Aspose.Cells embed PDF as OLE object | C# add OLE object to specific Excel cell | set OLE object dimensions with Aspose.Cells | how to set ProgID for embedded PDF in Excel | embed PDF file in Excel using Aspose.Cells .NET
// Developer Intent: Embed a PDF file as an OLE object in cell H4 of an Excel worksheet and control its size and ProgID using Aspose.Cells for .NET.
// Use Cases: Generate a report that shows a PDF preview directly inside the spreadsheet. | Create a self‑contained workbook that bundles related PDFs for offline distribution. | Automate the insertion of legal or specification documents into Excel templates.
// AI Prompts: Write C# code with Aspose.Cells to embed a PDF as an OLE object at cell D5 and set custom width and height. | Explain how to change the ProgID when embedding different file types (Word, Excel, PDF) with Aspose.Cells. | Provide a checklist for safely embedding external files as OLE objects in an Excel workbook using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // Required for OleObject

namespace Example
{
    // Demonstrates how to create a workbook, read a PDF file, add an OLE object placeholder at cell H4, set its width and height to 200 px, embed the PDF data, assign the PDF ProgID, and save the file with Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the PDF file to be embedded
                string pdfPath = "sample.pdf";

                // Verify the PDF file exists before attempting to read it
                if (!File.Exists(pdfPath))
                {
                    Console.WriteLine($"PDF file not found: {pdfPath}");
                    return;
                }

                // Read the PDF file as a byte array (the embedded object data)
                byte[] pdfData = File.ReadAllBytes(pdfPath);

                // Add an OLE object placeholder at cell H4 (row index 3, column index 7)
                // Height and width are set to 200 pixels each; imageData is an empty byte array
                int oleIndex = worksheet.OleObjects.Add(3, 7, 200, 200, new byte[0]);

                // Retrieve the newly added OleObject
                OleObject oleObject = worksheet.OleObjects[oleIndex];

                // Embed the PDF data into the OLE object
                // linkToFile = false (embed the file), objectData = pdfData,
                // sourceFileName = "sample.pdf", displayAsIcon = false,
                // label = "Embedded PDF"
                oleObject.SetEmbeddedObject(false, pdfData, "sample.pdf", false, "Embedded PDF");

                // Optionally set the ProgID for PDF files to improve handling
                oleObject.ProgID = "AcroExch.Document.DC";

                // Save the workbook with the embedded PDF OLE object
                string outputPath = "WorkbookWithPdfOleObject.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved: {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
