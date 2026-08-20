// Title: C# – Convert XLSX to PDF with Aspose.Cells (Load from File Path)
// Description: Loads an Excel workbook from a specified path, creates a minimal placeholder file when the source is missing, and saves the workbook as a PDF using Aspose.Cells' SaveFormat.Pdf.
// Keywords: Aspose.Cells | C# | XLSX to PDF | Workbook.Save | SaveFormat.Pdf | Excel to PDF conversion | load workbook from file | placeholder workbook
// Common Searches: Aspose.Cells convert xlsx to pdf c# | C# load excel file and save as pdf using Aspose | create placeholder workbook if file missing Aspose.Cells | save excel as pdf with Aspose.Cells .NET | batch convert excel files to pdf Aspose.Cells
// Developer Intent: Export an .xlsx workbook to a PDF document in a .NET application using Aspose.Cells.
// Use Cases: Generate PDF reports from existing Excel templates. | Automate bulk conversion of multiple XLSX files to PDF for archiving or distribution. | Provide a fallback placeholder workbook when the expected source file is absent, then export it to PDF.
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, creates a placeholder workbook if the file does not exist, and saves it as a PDF. | Explain step‑by‑step how to convert an Excel workbook to PDF using Aspose.Cells, including handling missing source files. | Show a C# example that batch processes a folder of .xlsx files, converting each to PDF with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;   // Aspose.Cells namespace provides Workbook and SaveFormat

namespace AsposeCellsConversionDemo
{
    // Loads an Excel workbook from a specified path, creates a minimal placeholder file when the source is missing, and saves the workbook as a PDF using Aspose.Cells' SaveFormat.Pdf.
    public class XlsxToPdf
    {
        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Path to the source XLSX file
            string sourcePath = "input.xlsx";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            // Ensure the source file exists; create a simple workbook if it does not
            if (!File.Exists(sourcePath))
            {
                var wb = new Workbook();
                var sheet = wb.Worksheets[0];
                sheet.Cells["A1"].PutValue("Sample Data");
                wb.Save(sourcePath, SaveFormat.Xlsx);
                Console.WriteLine($"Created placeholder workbook at '{sourcePath}'.");
            }

            // Load the workbook from the specified file
            Workbook workbook = new Workbook(sourcePath);

            // Save the loaded workbook as PDF
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Conversion completed: '{sourcePath}' -> '{pdfPath}'");
        }
    }
}
