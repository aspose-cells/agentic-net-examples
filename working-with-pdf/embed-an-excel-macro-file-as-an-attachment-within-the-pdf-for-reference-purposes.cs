// Title: Convert a macro‑enabled Excel (XLSM) workbook to PDF with Aspose.Cells for .NET (C#)
// Common Searches: C# code to save an XLSM workbook as a PDF using Aspose.Cells | Aspose.Cells .NET example for converting macro-enabled Excel to PDF | How to export an Excel file with macros to PDF in C# | Saving macro-enabled workbook to PDF with Aspose.Cells SaveFormat.Pdf | Aspose.Cells PDF conversion while preserving macro file reference
// Tags: Aspose.Cells XLSM to PDF conversion | C# export macro-enabled workbook to PDF | Aspose.Cells PDF save format usage | .NET convert Excel macro file to PDF | Aspose.Cells PDF export for Excel workbooks

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a macro‑enabled XLSM workbook with Aspose.Cells, verifies the file exists, and saves it as a PDF using the SaveFormat.Pdf option, handling any runtime exceptions.
class EmbedMacroInPdf
{
    static void Main()
    {
        try
        {
            // Path to the macro-enabled Excel file (XLSM)
            string macroFilePath = "macro.xlsm";

            // Verify that the macro file exists to avoid FileNotFoundException
            if (!File.Exists(macroFilePath))
            {
                Console.WriteLine($"Error: The file '{macroFilePath}' was not found.");
                return;
            }

            // Load the Excel workbook (contains macros)
            Workbook workbook = new Workbook(macroFilePath);

            // Define the output PDF file path
            string outputPdfPath = "output.pdf";

            // Save the workbook as PDF
            workbook.Save(outputPdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
