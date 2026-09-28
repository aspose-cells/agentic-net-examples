// Title: Automatically auto‑fit rows in every worksheet and export each sheet as a separate PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that opens an Excel workbook with Aspose.Cells, iterates through all worksheets, calls AutoFitRows on each sheet, and saves each sheet individually as a PDF with one page per sheet. | Create a C# routine that copies a worksheet into a new Workbook, configures PdfSaveOptions (OnePagePerSheet = true, AllColumnsInOnePagePerSheet = true), and writes the PDF to a uniquely named file.
// Common Searches: how to auto fit rows for each sheet before converting to PDF with Aspose.Cells .NET | save each worksheet of an Excel workbook as separate PDF files using Aspose.Cells C# | batch export Excel worksheets to PDF one page per sheet Aspose.Cells | move a single sheet into a separate workbook and generate a PDF using Aspose.Cells C#
// Tags: auto-fit rows Aspose.Cells C# | export worksheet to PDF Aspose.Cells | PdfSaveOptions OnePagePerSheet C# | temporary workbook copy sheet Aspose.Cells | batch process multiple worksheets Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program checks for the input Excel file, loads it with Aspose.Cells, loops through every worksheet, auto‑fits rows for proper rendering, copies each sheet into a temporary workbook, applies PdfSaveOptions to force one page per sheet and fit all columns, and saves each sheet as a uniquely named PDF while handling errors per sheet.
class BatchAutoFitRowsToPdf
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string inputPath = "input.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Iterate through each worksheet
            for (int sheetIndex = 0; sheetIndex < workbook.Worksheets.Count; sheetIndex++)
            {
                try
                {
                    Worksheet sheet = workbook.Worksheets[sheetIndex];

                    // Auto‑fit rows for proper rendering
                    sheet.AutoFitRows();

                    // Create a temporary workbook containing only the current sheet
                    Workbook tempWb = new Workbook();
                    tempWb.Worksheets.Clear();

                    // Copy the current sheet into the temporary workbook
                    tempWb.Worksheets.AddCopy(sheet.Name);

                    // Configure PDF save options
                    PdfSaveOptions pdfOptions = new PdfSaveOptions
                    {
                        OnePagePerSheet = true,
                        AllColumnsInOnePagePerSheet = true
                    };

                    // Output PDF file name
                    string outputPath = $"Sheet_{sheetIndex + 1}.pdf";

                    // Save the selected sheet as PDF
                    tempWb.Save(outputPath, pdfOptions);

                    Console.WriteLine($"Saved sheet {sheetIndex + 1} to \"{outputPath}\".");
                }
                catch (Exception exSheet)
                {
                    Console.WriteLine($"Error processing sheet {sheetIndex + 1}: {exSheet.Message}");
                }
            }

            Console.WriteLine("All worksheets have been auto‑fitted and saved as individual PDFs.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
