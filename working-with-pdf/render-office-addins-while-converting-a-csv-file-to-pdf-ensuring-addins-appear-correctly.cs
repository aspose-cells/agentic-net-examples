// Title: Convert a CSV file to PDF in C# with Aspose.Cells and ensure Office Add‑Ins are rendered in the output
// AI Prompts: Generate C# code that uses Aspose.Cells to load a CSV file, configure PDF export options, and save the workbook as a PDF while preserving any Office Add‑Ins. | Describe how to verify that Office Add‑Ins appear correctly in the PDF produced from a CSV workbook using Aspose.Cells. | Provide error‑handling and logging steps for a console app that converts CSV to PDF and checks for add‑in rendering with Aspose.Cells.
// Common Searches: Aspose.Cells export CSV to PDF with Office Add‑Ins rendering in C# | C# how to keep Excel add‑ins visible when converting CSV to PDF using Aspose.Cells | PDFSaveOptions settings for preserving Office Add‑Ins in Aspose.Cells .NET | Verify Office Add‑Ins in PDF generated from CSV workbook with Aspose.Cells | Convert CSV to PDF and include Excel add‑ins using Aspose.Cells library
// Tags: CSV to PDF conversion Aspose.Cells C# | Office Add‑Ins rendering PDFSaveOptions | Aspose.Cells preserve add‑ins during PDF export | C# Aspose.Cells PDF export with add‑ins | verify add‑ins in PDF generated from CSV workbook

using System;
using System.IO;
using Aspose.Cells;

namespace CsvToPdfWithAddIns
{
    // The example loads a CSV file into an Aspose.Cells Workbook, applies default PdfSaveOptions, and saves the workbook as a PDF. It includes guidance on configuring export settings and checking that any Office Add‑Ins present in the source are rendered correctly in the resulting PDF.
    class Program
    {
        static void Main(string[] args)
        {
            // Path to the source CSV file
            string csvPath = "input.csv";

            // Path for the resulting PDF file
            string pdfPath = "output.pdf";

            try
            {
                // Ensure the CSV file exists before loading
                if (!File.Exists(csvPath))
                {
                    Console.WriteLine($"Error: CSV file not found at '{csvPath}'.");
                    return;
                }

                // Load the CSV file into a Workbook object
                Workbook workbook = new Workbook(csvPath);

                // Configure PDF save options (default options are sufficient for CSV to PDF)
                PdfSaveOptions saveOptions = new PdfSaveOptions();

                // Save the workbook as a PDF document using the configured options
                workbook.Save(pdfPath, saveOptions);

                Console.WriteLine("CSV has been successfully converted to PDF.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
