// Title: Convert a TSV file to PDF with 0.5‑inch margins and fit‑to‑page width using Aspose.Cells for .NET
// AI Prompts: Generate C# code that reads a tab‑separated values (TSV) file into an Aspose.Cells Workbook, applies 0.5‑inch margins, and saves it as a PDF. | Show how to configure Aspose.Cells PageSetup to fit a worksheet to a single page width while preserving custom margins before exporting to PDF. | Provide error‑handling logic for a missing TSV input file and exception reporting in a C# Aspose.Cells PDF conversion routine.
// Common Searches: c# Aspose.Cells convert tsv to pdf with custom margins | how to set page margins in Aspose.Cells before PDF export | fit worksheet to one page width Aspose.Cells PDF output | Aspose.Cells load tab separated values file in .NET | adjust margins in points Aspose.Cells page setup
// Tags: TSV to PDF conversion Aspose.Cells | custom page margins Aspose.Cells | fit worksheet to single page width Aspose.Cells | LoadOptions LoadFormat.Tsv Aspose.Cells | margin values in points Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example verifies the TSV input file, loads it into an Aspose.Cells Workbook using LoadFormat.Tsv, accesses the first worksheet, sets 0.5‑inch margins (converted to points), configures the sheet to fit to one page width, and saves the workbook as a PDF while handling possible exceptions.
class TsvToPdfConverter
{
    static void Main()
    {
        // Input TSV file path
        string tsvPath = "input.tsv";

        // Output PDF file path
        string pdfPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(tsvPath))
        {
            Console.WriteLine($"Error: Input file not found at '{tsvPath}'.");
            return;
        }

        try
        {
            // Load the TSV file into a workbook using LoadFormat.Tsv
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Access the first worksheet (the TSV data is loaded here)
            Worksheet sheet = workbook.Worksheets[0];

            // Set custom page margins (values are in inches)
            // Aspose.Cells expects margin values in points; 1 inch = 72 points
            sheet.PageSetup.LeftMargin = 0.5 * 72;   // Left margin (0.5 inch)
            sheet.PageSetup.RightMargin = 0.5 * 72;  // Right margin (0.5 inch)
            sheet.PageSetup.TopMargin = 0.5 * 72;    // Top margin (0.5 inch)
            sheet.PageSetup.BottomMargin = 0.5 * 72; // Bottom margin (0.5 inch)

            // Optional: Fit the sheet to a single page width for better readability
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0; // Let height adjust automatically

            // Save the workbook as a PDF document
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"PDF successfully created at '{pdfPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
