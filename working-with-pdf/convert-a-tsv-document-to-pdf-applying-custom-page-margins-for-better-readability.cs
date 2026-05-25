using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTsvToPdf
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Paths for the source TSV file and the destination PDF file
                string tsvPath = "input.tsv";
                string pdfPath = "output.pdf";

                // Verify that the input TSV file exists
                if (!File.Exists(tsvPath))
                {
                    Console.WriteLine($"Error: The file \"{tsvPath}\" was not found.");
                    return;
                }

                // Load the TSV file into a workbook using LoadOptions to specify TSV format
                LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
                Workbook workbook = new Workbook(tsvPath, loadOptions);

                // Access the first worksheet (the TSV data is loaded into it)
                Worksheet sheet = workbook.Worksheets[0];

                // ----- Apply custom page margins (values in centimeters) -----
                sheet.PageSetup.TopMargin = 2.0;      // Top margin: 2 cm
                sheet.PageSetup.BottomMargin = 2.0;   // Bottom margin: 2 cm
                sheet.PageSetup.LeftMargin = 1.5;     // Left margin: 1.5 cm
                sheet.PageSetup.RightMargin = 1.5;    // Right margin: 1.5 cm

                // Optional: center the content horizontally and vertically
                sheet.PageSetup.CenterHorizontally = true;
                sheet.PageSetup.CenterVertically = true;

                // Create PDF save options; enable OnePagePerSheet to keep each sheet on a single page
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF using the specified options
                workbook.Save(pdfPath, pdfOptions);

                Console.WriteLine($"TSV file has been converted to PDF with custom margins: {pdfPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}