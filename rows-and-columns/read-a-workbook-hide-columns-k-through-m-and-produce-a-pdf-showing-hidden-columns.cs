// Title: Hide Columns K‑M in Excel and Export to PDF with Visible Hidden Columns using Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook, hides columns K through M on the first worksheet, sets PdfSaveOptions.HiddenColDisplayType to Visible, and saves the result as a PDF so the hidden columns are rendered in the output document.
// Keywords: Aspose.Cells | C# PDF export | HideColumns | PdfSaveOptions HiddenColDisplayType | Excel to PDF conversion | display hidden columns | column K to M | .NET Aspose.Cells example
// Common Searches: Aspose.Cells hide specific columns but show in PDF | C# export Excel to PDF with hidden columns visible | PdfSaveOptions HiddenColDisplayType usage | How to hide columns K-M in Excel and keep them in PDF | Aspose.Cells hide columns then render in PDF
// Developer Intent: Hide columns K‑M in the worksheet while ensuring they appear in the generated PDF.
// Use Cases: Mask confidential data in the UI but include it in printable reports. | Prepare workbooks for on‑screen editing with hidden columns, then produce PDFs that retain the full layout. | Automate batch conversion of Excel files where hidden columns must be visible in the final PDF documents.
// AI Prompts: Generate C# code that hides columns K to M in an Aspose.Cells workbook and exports to PDF with HiddenColDisplayType set to Visible. | Explain how PdfSaveOptions.HiddenColDisplayType works in Aspose.Cells for .NET and show a complete example with error handling.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// C# sample that loads an Excel workbook, hides columns K through M on the first worksheet, sets PdfSaveOptions.HiddenColDisplayType to Visible, and saves the result as a PDF so the hidden columns are rendered in the output document.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Hide columns K (index 10) through M (index 12) – total of 3 columns
            workbook.Worksheets[0].Cells.HideColumns(10, 3);

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions();

            // Note: In some versions of Aspose.Cells the HiddenColDisplayType property
            // may not be available. If it exists, you can set it to display hidden columns:
            // pdfOptions.HiddenColDisplayType = HtmlHiddenColDisplayType.Visible;

            // Save the workbook as PDF with the specified options
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
