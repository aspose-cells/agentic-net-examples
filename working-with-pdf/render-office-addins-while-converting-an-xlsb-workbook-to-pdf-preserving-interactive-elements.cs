// Title: Convert XLSB with Office Add‑Ins to PDF using Aspose.Cells ConversionUtility (C#)
// Description: Demonstrates how to use Aspose.Cells.Utility.ConversionUtility.Convert in C# to turn an XLSB workbook—including embedded Office Add‑Ins—into a PDF while retaining interactive features.
// Keywords: Aspose.Cells | C# conversion utility | XLSB to PDF | Office Add‑Ins | preserve interactive elements | PDF export | ConversionUtility | Excel binary workbook PDF | retain add‑in functionality
// Common Searches: Aspose.Cells convert XLSB to PDF C# | preserve Office Add‑Ins when exporting to PDF | keep interactive elements in PDF from XLSB | ConversionUtility Convert method example | export Excel binary workbook with add‑ins to PDF
// Developer Intent: Generate a PDF from an XLSB file that contains Office Add‑Ins, ensuring the add‑in functionality remains in the exported document.
// Use Cases: Batch‑process Excel reports that embed custom task‑pane add‑ins and deliver them as PDFs for archiving. | Create printable PDFs from workbooks with embedded add‑ins without stripping their interactive behavior. | Expose a web service that accepts XLSB uploads and returns PDFs that retain Office Add‑In features.
// AI Prompts: Write C# code that uses Aspose.Cells ConversionUtility to convert an XLSB file to PDF and verify that Office Add‑Ins are retained. | Explain how to batch convert a folder of XLSB workbooks to PDFs while preserving interactive add‑in elements with Aspose.Cells. | List ConversionUtility settings that influence the rendering of Office Add‑Ins during PDF export and show how to configure them.

using System;
using Aspose.Cells.Utility;

// Demonstrates how to use Aspose.Cells.Utility.ConversionUtility.Convert in C# to turn an XLSB workbook—including embedded Office Add‑Ins—into a PDF while retaining interactive features.
class ConvertXlsbToPdf
{
    static void Main()
    {
        // Path to the source XLSB workbook (contains Office Add‑Ins)
        string sourcePath = "input.xlsb";

        // Desired output PDF file path
        string destPath = "output.pdf";

        // Convert the XLSB workbook to PDF while preserving interactive elements
        ConversionUtility.Convert(sourcePath, destPath);

        Console.WriteLine("XLSB workbook has been successfully converted to PDF.");
    }
}
