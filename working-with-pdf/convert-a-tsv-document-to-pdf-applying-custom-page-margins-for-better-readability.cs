// Title: Convert a TSV spreadsheet to PDF with custom margins using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a .tsv file into an Aspose.Cells Workbook, sets top, bottom, left, and right margins, and exports the first worksheet to a PDF. | Show how to apply 0.5‑inch top/bottom and 0.75‑inch left/right margins and fit the worksheet to one page wide when saving a workbook as PDF with Aspose.Cells. | Write a console application that validates the input TSV path, creates the output directory if missing, and performs the conversion to PDF with custom page setup using LoadOptions and PdfSaveOptions.
// Common Searches: C# Aspose.Cells set page margins before exporting TSV to PDF | How to fit worksheet to one page width when saving TSV as PDF with Aspose.Cells | Load TSV file into Aspose.Cells workbook and customize PDF margins | Aspose.Cells PDFSaveOptions margin settings for TSV conversion in .NET
// Tags: TSV to PDF Aspose.Cells | PageSetup margin settings Aspose.Cells | FitToPagesWide usage Aspose.Cells | LoadOptions TSV Aspose.Cells | PdfSaveOptions configuration Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads a TSV file into an Aspose.Cells Workbook, configures the worksheet’s PageSetup margins (0.5 in top/bottom, 0.75 in left/right), sets FitToPagesWide = 1, and saves the sheet as a PDF using PdfSaveOptions, while handling missing input files and creating the output folder.
class TsvToPdfConverter
{
    static void Main()
    {
        // Paths for input TSV and output PDF
        string tsvPath = @"C:\Input\data.tsv";
        string pdfPath = @"C:\Output\data.pdf";

        try
        {
            // Verify that the source TSV file exists
            if (!File.Exists(tsvPath))
            {
                Console.WriteLine($"Input file not found: {tsvPath}");
                return;
            }

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(pdfPath);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            // Load the TSV file into a Workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom page margins (values are in points; 1 inch = 72 points)
            sheet.PageSetup.TopMargin = 0.5 * 72;    // 0.5 inch
            sheet.PageSetup.BottomMargin = 0.5 * 72; // 0.5 inch
            sheet.PageSetup.LeftMargin = 0.75 * 72;  // 0.75 inch
            sheet.PageSetup.RightMargin = 0.75 * 72; // 0.75 inch

            // Optional: Fit the sheet to a single page width for better readability
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0; // 0 means unlimited height

            // Save the workbook as a PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                AllColumnsInOnePagePerSheet = false
            };
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("TSV has been successfully converted to PDF with custom margins.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
