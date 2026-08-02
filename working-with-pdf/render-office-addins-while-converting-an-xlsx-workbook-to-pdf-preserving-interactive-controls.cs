// Title: C# – Convert an XLSX workbook with Office Add‑Ins to PDF while retaining interactive elements using Aspose.Cells
// Description: Shows how to load an Excel file that contains Office Add‑Ins (checkboxes, drop‑downs, ActiveX controls) and export it to PDF with those controls still visible, using Aspose.Cells ConversionUtility.Convert together with file‑existence checks and exception handling.
// Keywords: Aspose.Cells | C# | XLSX to PDF | Office Add‑Ins | interactive elements | ConversionUtility | Excel form controls | PDF export | automation | batch conversion
// Common Searches: Aspose.Cells preserve Office Add‑Ins in PDF | convert Excel with form controls to PDF C# | keep interactive elements when exporting to PDF Aspose | ConversionUtility retain add‑ins PDF | Excel dashboard PDF with checkboxes Aspose.Cells
// Developer Intent: Export an Excel workbook that includes Office Add‑Ins to a PDF without losing the interactive controls.
// Use Cases: Create printable PDFs of Excel dashboards that contain checkboxes, drop‑downs, or other form controls while keeping their visual representation. | Archive compliance‑critical Excel reports with embedded add‑ins, ensuring the PDF version still displays the interactive elements. | Automate batch processing of multiple workbooks that contain Office Add‑Ins, producing PDFs ready for stakeholder distribution.
// AI Prompts: Write C# code that uses Aspose.Cells ConversionUtility to convert an XLSX file with Office Add‑Ins to PDF while keeping the controls intact. | Explain how Aspose.Cells handles Excel form controls and ActiveX objects during PDF export and which options affect their preservation. | Show how to add robust error handling and file‑existence validation when converting workbooks that contain add‑ins to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;

namespace AsposeCellsAddInPdfConversion
{
    // Shows how to load an Excel file that contains Office Add‑Ins (checkboxes, drop‑downs, ActiveX controls) and export it to PDF with those controls still visible, using Aspose.Cells ConversionUtility.Convert together with file‑existence checks and exception handling.
    class Program
    {
        static void Main()
        {
            // Path to the source XLSX workbook that contains Office Add‑Ins (interactive controls)
            string sourcePath = "input_with_addins.xlsx";

            // Desired output PDF file path
            string outputPath = "output_preserving_addins.pdf";

            try
            {
                // Verify that the source file exists to avoid FileNotFoundException
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Convert the workbook to PDF while preserving interactive controls.
                // ConversionUtility handles loading and saving internally.
                ConversionUtility.Convert(sourcePath, outputPath);

                Console.WriteLine($"Workbook converted successfully to PDF: {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any runtime errors and display a friendly message
                Console.WriteLine($"Error during conversion: {ex.Message}");
            }
        }
    }
}
