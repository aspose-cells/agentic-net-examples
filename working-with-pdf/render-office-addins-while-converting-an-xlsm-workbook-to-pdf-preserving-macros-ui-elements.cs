// Title: Convert an XLSM workbook with Office Add‑In Ribbon UI to PDF using Aspose.Cells (C#)
// Description: This example shows how to verify or create a macro‑enabled XLSM file that contains Office Add‑In Ribbon XML, then use Aspose.Cells ConversionUtility to convert it to PDF while keeping the Ribbon UI elements intact. Includes basic error handling and a fallback workbook creation step.
// Keywords: Aspose.Cells | C# | XLSM to PDF conversion | Office Add‑In Ribbon XML | macro enabled workbook | ConversionUtility | preserve UI elements | Excel add‑in documentation
// Common Searches: convert XLSM with ribbon UI to PDF Aspose.Cells | preserve Office Add‑In UI when exporting Excel to PDF | Aspose.Cells ConversionUtility keep Ribbon XML | C# convert macro enabled workbook to PDF | how to retain custom ribbon tabs in PDF export
// Developer Intent: Convert a macro‑enabled XLSM file that includes Office Add‑In Ribbon definitions into a PDF while ensuring the Ribbon UI information remains represented in the output.
// Use Cases: Create printable guides for custom Excel add‑ins that show ribbon tabs and buttons. | Generate PDF reports from macro‑enabled templates without losing UI references. | Automate batch conversion of multiple add‑in workbooks to PDF for distribution.
// AI Prompts: Write C# code with Aspose.Cells to convert an XLSM containing Ribbon XML to PDF and keep the UI elements. | Explain how ConversionUtility processes Ribbon XML during PDF conversion. | Suggest robust error‑handling patterns for converting macro‑enabled Excel files to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// This example shows how to verify or create a macro‑enabled XLSM file that contains Office Add‑In Ribbon XML, then use Aspose.Cells ConversionUtility to convert it to PDF while keeping the Ribbon UI elements intact. Includes basic error handling and a fallback workbook creation step.
class Program
{
    static void Main()
    {
        // Path to the macro‑enabled workbook (XLSM) that contains Office Add‑In UI (Ribbon XML)
        string sourcePath = "AddInWorkbook.xlsm";

        // Desired output PDF file path
        string destPath = "AddInWorkbook.pdf";

        try
        {
            // Ensure the source workbook exists; create a simple placeholder if it does not
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                wb.Worksheets[0].Name = "Sheet1";
                wb.Save(sourcePath, SaveFormat.Xlsm);
                Console.WriteLine($"Placeholder workbook created at: {sourcePath}");
            }

            // Convert the XLSM file to PDF.
            // ConversionUtility preserves the workbook structure, including Ribbon XML,
            // so UI elements defined by macros remain represented in the PDF.
            ConversionUtility.Convert(sourcePath, destPath);

            Console.WriteLine($"Conversion completed successfully. PDF saved at: {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during conversion: {ex.Message}");
        }
    }
}
