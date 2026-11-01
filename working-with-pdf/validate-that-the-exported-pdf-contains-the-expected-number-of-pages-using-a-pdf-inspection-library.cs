// Title: How to verify the page count of a PDF generated from an Excel workbook using Aspose.Cells in C#
// Common Searches: C# get number of pages from Excel to PDF conversion with Aspose.Cells | Aspose.Cells WorkbookRender page count before saving PDF | Validate expected PDF page count after exporting workbook in .NET | Check if generated PDF has correct number of pages using Aspose.Cells API | Compare actual and expected PDF pages after saving Excel as PDF in C#
// Tags: Aspose.Cells WorkbookRender page count | Excel to PDF page verification C# | Aspose.Cells PDF export page count validation | C# PDF generation from workbook with page count check | ImageOrPrintOptions PDF rendering Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The example loads an Excel file, creates a WorkbookRender with ImageOrPrintOptions to obtain the workbook's page count, saves the workbook as a PDF, and then compares the retrieved page count to an expected value, reporting success or failure and confirming the PDF file was created.
class Program
{
    static void Main()
    {
        try
        {
            // Input Excel file that will be exported to PDF
            string excelPath = "input.xlsx";

            // Verify that the Excel file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: The file \"{excelPath}\" was not found.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Define PDF output path
            string pdfPath = "output.pdf";

            // Set rendering options (required for page count calculation)
            ImageOrPrintOptions printOptions = new ImageOrPrintOptions
            {
                // Although SaveFormat is obsolete for image rendering, it does not affect PDF saving
                SaveFormat = SaveFormat.Pdf
            };

            // Create renderer to obtain page count
            WorkbookRender renderer = new WorkbookRender(workbook, printOptions);
            int actualPageCount = renderer.PageCount;

            // Export workbook to PDF using the proper API
            workbook.Save(pdfPath, SaveFormat.Pdf);

            // Expected number of pages (adjust as needed)
            int expectedPageCount = 3;

            // Validate page count
            if (actualPageCount == expectedPageCount)
            {
                Console.WriteLine("PDF page count validation passed.");
            }
            else
            {
                Console.WriteLine($"PDF page count validation failed. Expected {expectedPageCount}, but found {actualPageCount}.");
            }

            // Confirm that the PDF file was created
            if (File.Exists(pdfPath))
            {
                Console.WriteLine($"PDF file generated at: {Path.GetFullPath(pdfPath)}");
            }
            else
            {
                Console.WriteLine("Error: PDF file was not created.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
