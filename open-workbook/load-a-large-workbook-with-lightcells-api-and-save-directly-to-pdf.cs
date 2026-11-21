// Title: Convert a large XLSX workbook to PDF in C# with Aspose.Cells using LoadOptions for memory‑efficient processing
// Common Searches: C# Aspose.Cells load large Excel file with LoadOptions and save as PDF without out‑of‑memory errors | how to use Aspose.Cells LoadOptions to open a big .xlsx and export to PDF in .NET | Aspose.Cells convert massive workbook to PDF directly in C# example | memory‑optimized Excel to PDF conversion using Aspose.Cells in a console application | save large workbook to PDF using Aspose.Cells without intermediate files
// Tags: load large xlsx with Aspose.Cells LoadOptions | export workbook to PDF using Aspose.Cells | memory efficient Excel to PDF conversion .NET | Aspose.Cells C# workbook to PDF example | large workbook processing Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for the source XLSX file, creates LoadOptions for the Xlsx format, loads the workbook with Aspose.Cells, and saves it directly to PDF, handling exceptions and reporting the result.
class Program
{
    static void Main()
    {
        // Path to the source workbook (large file)
        string sourcePath = "LargeWorkbook.xlsx";

        // Path for the resulting PDF
        string pdfPath = "LargeWorkbook.pdf";

        try
        {
            // Ensure the source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load options (specify the format if needed)
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);

            // Load the workbook (standard loading; suitable for most scenarios)
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Save the workbook directly to PDF format
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Conversion successful. PDF saved at: {pdfPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors gracefully
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
