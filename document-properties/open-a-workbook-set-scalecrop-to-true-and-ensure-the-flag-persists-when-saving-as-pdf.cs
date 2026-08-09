// Title: C# – Set Workbook ScaleCrop flag and retain it when exporting to PDF with Aspose.Cells
// Description: Demonstrates how to enable the BuiltInDocumentProperties.ScaleCrop flag on an Aspose.Cells workbook, verify the setting, and save the workbook as a PDF so the flag is preserved in the output file.
// Keywords: Aspose.Cells ScaleCrop | C# ScaleCrop PDF | BuiltInDocumentProperties ScaleCrop | preserve workbook properties PDF | Aspose.Cells PDF export | thumbnail scaling Excel PDF
// Common Searches: how to set scalecrop in Aspose.Cells C# | Aspose.Cells keep ScaleCrop flag when saving PDF | ScaleCrop property PDF export Aspose.Cells | C# enable workbook thumbnail scaling Aspose | Aspose.Cells preserve document properties in PDF
// Developer Intent: Enable the ScaleCrop property on a workbook and ensure it remains set after the workbook is saved as a PDF.
// Use Cases: Create a new workbook, turn on ScaleCrop, and generate a PDF with a scaled thumbnail. | Load an existing Excel file, modify its ScaleCrop flag, and export to PDF while keeping the flag for downstream viewers. | Automate report generation where PDF thumbnails must follow the ScaleCrop setting defined in the source workbook.
// AI Prompts: Show C# code that sets BuiltInDocumentProperties.ScaleCrop to true on an Aspose.Cells workbook and saves it as a PDF preserving the flag. | Provide an example of loading an Excel file with Aspose.Cells, enabling ScaleCrop, and exporting to PDF with the property retained. | Explain whether any PdfSaveOptions settings are required to keep the ScaleCrop flag when converting a workbook to PDF using Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Properties;
using Aspose.Cells.Rendering;

// Demonstrates how to enable the BuiltInDocumentProperties.ScaleCrop flag on an Aspose.Cells workbook, verify the setting, and save the workbook as a PDF so the flag is preserved in the output file.
class Program
{
    static void Main()
    {
        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Set the ScaleCrop property to true so the thumbnail is displayed in scaled mode
        workbook.BuiltInDocumentProperties.ScaleCrop = true;

        // Verify that the property is set
        Console.WriteLine("ScaleCrop property value: " + workbook.BuiltInDocumentProperties.ScaleCrop);

        // Prepare PDF save options (default options are sufficient for preserving the property)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as PDF; the ScaleCrop flag will be persisted in the output file
        workbook.Save("output.pdf", pdfOptions);
    }
}
