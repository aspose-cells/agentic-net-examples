// Title: Convert a TSV file to a multi‑page PDF with one page per worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a .tsv file into an Aspose.Cells Workbook, deletes any completely empty worksheets, and saves the workbook as a PDF where each worksheet appears on its own page. | Show how to set PdfSaveOptions in Aspose.Cells to enable OnePagePerSheet and avoid generating blank pages when converting from TSV. | Add comprehensive error handling that checks for the presence of the input TSV file and catches exceptions during loading or PDF saving.
// Common Searches: aspnet convert tsv to pdf using aspose.cells one page per sheet | c# remove empty worksheets before saving pdf with aspose.cells | how to prevent blank pages in pdf generated from excel workbook aspose.cells | load tsv file into workbook with loadoptions aspose.cells c#
// Tags: Aspose.Cells TSV import | PdfSaveOptions OnePagePerSheet | remove empty worksheets Aspose.Cells | prevent blank PDF pages Aspose.Cells | C# workbook to PDF conversion Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// This C# example loads a TSV file into an Aspose.Cells Workbook, removes any completely empty worksheets to stop blank PDF pages, configures PdfSaveOptions for one page per sheet, and saves the result as a PDF with robust file‑existence checks and exception handling.
class Program
{
    static void Main()
    {
        // Input TSV file and output PDF file paths
        string tsvPath = "input.tsv";
        string pdfPath = "output.pdf";

        // Verify that the input file exists
        if (!File.Exists(tsvPath))
        {
            Console.WriteLine($"Input file not found: {tsvPath}");
            return;
        }

        try
        {
            // Load the TSV workbook
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Tsv);
            Workbook workbook = new Workbook(tsvPath, loadOptions);

            // Remove completely empty worksheets to prevent blank PDF pages
            for (int i = workbook.Worksheets.Count - 1; i >= 0; i--)
            {
                Worksheet sheet = workbook.Worksheets[i];
                if (sheet.Cells.MaxDataRow < 0 || sheet.Cells.MaxDataColumn < 0)
                {
                    workbook.Worksheets.RemoveAt(i);
                }
            }

            // Configure PDF save options: one page per sheet
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = true
                // IgnorePrintAreas property is not available in this version of Aspose.Cells
            };

            // Save the workbook as a PDF
            workbook.Save(pdfPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to: {pdfPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
