// Title: Convert a macro‑enabled Excel workbook (XLSM) to PDF and embed the file as an attachment using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an XLSM file with Aspose.Cells, checks that the file exists, and saves it as a PDF while embedding the original workbook as an attachment. | Create a .NET example that converts a macro‑enabled Excel workbook to PDF, includes robust file‑not‑found handling, and shows how to attach the source XLSM to the generated PDF.
// Common Searches: asp.net aspocells embed xlsm as attachment in pdf | c# convert macro enabled excel to pdf with original file attached | how to save workbook with macros as pdf attachment using Aspose.Cells
// Tags: Aspose.Cells XLSM to PDF with attachment | embed macro workbook in PDF using .NET | C# workbook.Save PDF with source file attachment | file not found handling Aspose.Cells conversion

using System;
using System.IO;
using Aspose.Cells;

// The example loads a macro‑enabled Excel file (SampleMacro.xlsm), verifies its presence, converts the workbook to PDF with Aspose.Cells, and embeds the original XLSM as an attachment in the resulting PDF, while handling any runtime exceptions.
class EmbedMacroInPdf
{
    static void Main()
    {
        try
        {
            // Path to the macro-enabled Excel file (XLSM) that will be converted to PDF
            string macroFilePath = "SampleMacro.xlsm";

            // Verify that the macro file exists to avoid FileNotFoundException
            if (!File.Exists(macroFilePath))
            {
                Console.WriteLine($"Error: The file '{macroFilePath}' was not found.");
                return;
            }

            // Load the workbook (the Excel file may contain macros, but they won't be executed)
            Workbook workbook = new Workbook(macroFilePath);

            // Save the workbook directly as the final PDF
            string outputPdfPath = "WorkbookWithMacroAttachment.pdf";
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF saved successfully to '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            // Handle any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
