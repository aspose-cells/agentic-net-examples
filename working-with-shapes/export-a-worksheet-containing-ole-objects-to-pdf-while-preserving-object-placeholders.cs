// Title: Export OLE Object Placeholders to PDF with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to add an embedded OLE object displayed as an icon, configure the workbook to show placeholders instead of the actual objects, and save the worksheet as a PDF using Aspose.Cells. The PDF contains only the OLE icons, does not embed the original files, and includes document structure for accessibility.
// Keywords: Aspose.Cells | C# PDF export | OLE object placeholder | DisplayDrawingObjects.Placeholders | PdfSaveOptions EmbedAttachments false | Export OLE icons to PDF | OleObjects.Add C# | Worksheet OLE placeholder PDF | Aspose.Cells PDF save options
// Common Searches: Aspose.Cells export OLE icons as placeholders PDF | C# display OLE objects as icons when saving to PDF | Prevent OLE attachments in PDF with Aspose.Cells | Set DisplayDrawingObjects to Placeholders before PDF conversion | PdfSaveOptions ExportDocumentStructure example
// Developer Intent: Generate a PDF from an Excel worksheet that contains OLE objects, showing only their icon placeholders and excluding embedded data.
// Use Cases: Secure reporting: share a PDF where embedded files are represented by icons, preventing distribution of the original content. | Template documentation: export a workbook with OLE placeholders for training manuals while keeping file size minimal. | Visual review: produce a PDF that preserves the layout of OLE icons for stakeholder review without exposing the embedded files.
// AI Prompts: Write C# code that inserts an OLE object, sets it to display as an icon, and saves the workbook to PDF with only placeholders using Aspose.Cells. | Explain how DisplayDrawingObjects.Placeholders and PdfSaveOptions.EmbedAttachments affect the PDF output of OLE objects in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add an embedded OLE object displayed as an icon, configure the workbook to show placeholders instead of the actual objects, and save the worksheet as a PDF using Aspose.Cells. The PDF contains only the OLE icons, does not embed the original files, and includes document structure for accessibility.
class ExportOleObjectsToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add some sample text
        worksheet.Cells["A1"].PutValue("Worksheet with OLE object placeholders");

        // -----------------------------------------------------------------
        // Add an OLE object (embedded) to the worksheet
        // -----------------------------------------------------------------
        // Image that will be shown as the OLE object's icon (placeholder)
        byte[] iconImage = File.ReadAllBytes("icon.png"); // replace with a real image file

        // Binary data of the embedded file (e.g., another Excel file)
        byte[] oleData = File.ReadAllBytes("sample.xlsx"); // replace with a real file

        // Add the OLE object using the Add method (topRow, leftColumn, height, width, imageData)
        int oleIndex = worksheet.OleObjects.Add(5, 2, 200, 200, iconImage);
        OleObject oleObject = worksheet.OleObjects[oleIndex];

        // Set the embedded object data and display options
        oleObject.ObjectData = oleData;
        oleObject.DisplayAsIcon = true;          // show as an icon
        oleObject.Label = "Sample Excel File";   // icon label

        // -----------------------------------------------------------------
        // Configure the workbook to show placeholders instead of actual objects
        // -----------------------------------------------------------------
        workbook.Settings.DisplayDrawingObjects = DisplayDrawingObjects.Placeholders;

        // -----------------------------------------------------------------
        // Set PDF save options
        // -----------------------------------------------------------------
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Ensure that OLE objects are not embedded as attachments (placeholders only)
        pdfOptions.EmbedAttachments = false; // default, but set explicitly per rule

        // Optional: export document structure for better accessibility
        pdfOptions.ExportDocumentStructure = true; // per rule

        // -----------------------------------------------------------------
        // Save the workbook as PDF
        // -----------------------------------------------------------------
        workbook.Save("WorksheetWithOlePlaceholders.pdf", pdfOptions);
    }
}
