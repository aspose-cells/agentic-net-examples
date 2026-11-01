// Title: Generate a landscape-oriented PDF from a wide Excel worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, sets the first worksheet's PageSetup.Orientation to Landscape, configures FitToPagesWide = 1, and saves the workbook as a PDF/A‑1b compliant document with Aspose.Cells. | Create a C# example that builds a sample workbook with many columns, applies landscape page orientation, fits the sheet width to a single PDF page, and uses PdfSaveOptions to produce the PDF. | Show how to verify the output folder exists, handle exceptions, and convert a wide Excel sheet to a landscape PDF while preserving PDF/A‑1b compliance in Aspose.Cells.
// Common Searches: aspnet convert wide excel sheet to landscape pdf using aspose.cells | c# set page orientation landscape and fit to one page wide when saving excel as pdf | how to generate pdf/a-1b from excel workbook with landscape layout in .net | aspose.cells fit worksheet width to single pdf page landscape orientation | sample code for exporting wide excel worksheet to landscape pdf in c#
// Tags: Aspose.Cells PDF export landscape orientation | fit worksheet width to one PDF page | C# Aspose.Cells PdfSaveOptions PDF/A-1b | convert wide Excel sheet to PDF | page setup orientation landscape Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program checks for an input Excel file (creating a sample wide worksheet if missing), loads it with Aspose.Cells, sets the first worksheet's page orientation to Landscape, configures the page setup to fit the sheet width to a single PDF page while allowing multiple pages in height, applies PdfSaveOptions with PDF/A‑1b compliance, ensures the output directory exists, and saves the workbook as a landscape‑oriented PDF.
class ExcelToLandscapePdf
{
    static void Main()
    {
        // Paths for input Excel and output PDF
        string excelPath = @"C:\Input\WideSheet.xlsx";
        string pdfPath = @"C:\Output\WideSheet_Landscape.pdf";

        try
        {
            // Ensure the input file exists; create a simple workbook if it does not.
            if (!File.Exists(excelPath))
            {
                // Create directory if needed
                Directory.CreateDirectory(Path.GetDirectoryName(excelPath));

                // Generate a sample workbook with wide data
                Workbook sampleWb = new Workbook();
                Worksheet ws = sampleWb.Worksheets[0];
                ws.Name = "WideData";

                // Populate many columns to simulate a wide sheet
                for (int col = 0; col < 30; col++)
                {
                    ws.Cells[0, col].PutValue($"Header {col + 1}");
                    for (int row = 1; row <= 20; row++)
                    {
                        ws.Cells[row, col].PutValue($"R{row}C{col + 1}");
                    }
                }

                // Save the sample workbook
                sampleWb.Save(excelPath);
            }

            // Load the workbook from the Excel file
            Workbook workbook = new Workbook(excelPath);

            // Work with the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Set page orientation to Landscape for wide sheets
            sheet.PageSetup.Orientation = PageOrientationType.Landscape;

            // Fit to one page wide; height can span multiple pages
            sheet.PageSetup.FitToPagesWide = 1;
            sheet.PageSetup.FitToPagesTall = 0;

            // Configure PDF save options
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                OnePagePerSheet = false,               // Keep default paging per sheet
                Compliance = PdfCompliance.PdfA1b      // Preserve layout compliance
            };

            // Ensure output directory exists
            Directory.CreateDirectory(Path.GetDirectoryName(pdfPath));

            // Save the workbook as a PDF
            workbook.Save(pdfPath, pdfOptions);

            Console.WriteLine("PDF generated successfully in landscape orientation.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
