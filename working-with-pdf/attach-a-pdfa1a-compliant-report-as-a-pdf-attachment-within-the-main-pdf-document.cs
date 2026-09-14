// Title: Attach a PDF/A‑1a compliant report as an OLE object to an Excel workbook using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx file with Aspose.Cells, reads a PDF/A‑1a file into a byte array, embeds it as an OLE object on the first worksheet, sets the object name and enables auto‑size, then saves the updated workbook. | Show how to verify the existence of the Excel and PDF files, handle possible exceptions, and configure OleObject properties (Name, IsAutoSize) when attaching a PDF report to a workbook with Aspose.Cells.
// Common Searches: c# aspose.cells embed pdf/a-1a report as ole object in excel workbook | how to add pdf attachment to existing .xlsx using aspose.cells library | asp.net core load pdf file into byte array and insert as ole object in worksheet | aspose.cells example for embedding external pdf into first worksheet | error handling when embedding pdf as ole object with aspose.cells c#
// Tags: pdf oleobject insertion aspose.cells c# | excel workbook pdf attachment aspose | configure oleobject name autosize aspose.cells | read pdf into byte array aspose.cells | save workbook with embedded pdf aspose

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing; // For OleObject

// The example checks that both the main Excel file and the PDF/A‑1a report exist, loads the workbook with Aspose.Cells, reads the PDF into a byte array, adds it as an OLE object to the first worksheet, sets the object's name and auto‑size property, and saves the workbook containing the embedded PDF.
class PdfAttachmentExample
{
    static void Main()
    {
        try
        {
            const string mainFilePath = "MainDocument.xlsx";
            const string reportFilePath = "Report.pdf";
            const string outputFilePath = "MainDocument_WithReport.xlsx";

            // Verify that the main Excel file exists
            if (!File.Exists(mainFilePath))
                throw new FileNotFoundException($"The main Excel file '{mainFilePath}' was not found.");

            // Verify that the PDF report exists
            if (!File.Exists(reportFilePath))
                throw new FileNotFoundException($"The PDF report file '{reportFilePath}' was not found.");

            // Load the main workbook
            Workbook workbook = new Workbook(mainFilePath);

            // Ensure there is at least one worksheet
            if (workbook.Worksheets.Count == 0)
                throw new InvalidOperationException("The workbook does not contain any worksheets.");

            // Read PDF file into a byte array (required by OleObjects.Add)
            byte[] pdfData = File.ReadAllBytes(reportFilePath);

            // Add the PDF as an OLE object to the first worksheet
            // Parameters: upper-left row, column, height (pixels), width (pixels), byte[] data
            int oleIndex = workbook.Worksheets[0].OleObjects.Add(0, 0, 100, 100, pdfData);

            // Configure the OLE object (optional)
            OleObject ole = workbook.Worksheets[0].OleObjects[oleIndex];
            ole.Name = "Report.pdf";
            ole.IsAutoSize = true; // Auto‑size to fit the cell

            // Save the updated workbook with the embedded PDF
            workbook.Save(outputFilePath);
            Console.WriteLine($"Workbook saved successfully to '{outputFilePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
