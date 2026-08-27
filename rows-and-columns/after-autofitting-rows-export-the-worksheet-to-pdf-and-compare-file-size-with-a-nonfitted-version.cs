// Title: Compare PDF file sizes before and after using Worksheet.AutoFitRows in Aspose.Cells (C#)
// AI Prompts: Write C# code that creates a workbook, saves it as a PDF, then invokes Worksheet.AutoFitRows, saves a second PDF, and prints the byte sizes of both files. | Show how to configure PdfSaveOptions and calculate the file length of PDFs generated with and without row auto‑fitting using Aspose.Cells. | Provide a C# example that wraps cell text, auto‑fits rows, exports the worksheet to PDF, and logs whether the PDF size increased, decreased, or stayed the same.
// Common Searches: how to measure PDF size difference after applying Worksheet.AutoFitRows in Aspose.Cells C# | Aspose.Cells export worksheet to PDF before and after auto fitting rows | does auto fitting rows affect PDF output size in Aspose.Cells .NET | C# code sample for comparing PDF file sizes with and without row auto fit using Aspose.Cells | PdfSaveOptions without auto fit rows Aspose.Cells size comparison
// Tags: auto-fit rows PDF size comparison | Aspose.Cells PdfSaveOptions row height | C# measure PDF file size Aspose.Cells | row auto fitting impact on PDF output | export worksheet to PDF Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsExamples
{
    // The example creates a workbook, adds wrapped text to cells, saves a PDF without auto‑fitting rows, records its size, calls sheet.AutoFitRows(), saves a second PDF, records the new size, and prints a comparison indicating whether the auto‑fit operation increased, decreased, or left the PDF file size unchanged.
    public class AutoFitRowsPdfComparison
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and access the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data that will cause varying row heights
                sheet.Cells["A1"].PutValue("Short text");
                sheet.Cells["A2"].PutValue("This is a longer text that should increase the row height when wrapped.");
                sheet.Cells["A3"].PutValue("Another long text\nwith line breaks\nto demonstrate auto‑fit behavior.");

                // Enable text wrapping for the cells to allow row height changes
                for (int row = 0; row < 3; row++)
                {
                    Style style = sheet.Cells[row, 0].GetStyle();
                    style.IsTextWrapped = true;
                    sheet.Cells[row, 0].SetStyle(style);
                }

                // Define file paths for the PDFs
                string basePath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                // Ensure the directory exists
                if (!Directory.Exists(basePath))
                {
                    Directory.CreateDirectory(basePath);
                }

                string pdfWithoutFit = Path.Combine(basePath, "Worksheet_NoAutoFit.pdf");
                string pdfWithFit = Path.Combine(basePath, "Worksheet_AutoFit.pdf");

                // Save PDF without auto‑fitting rows
                PdfSaveOptions pdfOptions = new PdfSaveOptions();
                workbook.Save(pdfWithoutFit, pdfOptions);

                // Get file size of the non‑fitted PDF
                long sizeWithoutFit = new FileInfo(pdfWithoutFit).Length;

                // Auto‑fit all rows in the worksheet
                sheet.AutoFitRows();

                // Save PDF after auto‑fitting rows
                workbook.Save(pdfWithFit, pdfOptions);

                // Get file size of the fitted PDF
                long sizeWithFit = new FileInfo(pdfWithFit).Length;

                // Output the comparison results
                Console.WriteLine($"PDF size without AutoFitRows: {sizeWithoutFit} bytes");
                Console.WriteLine($"PDF size with AutoFitRows   : {sizeWithFit} bytes");
                if (sizeWithFit > sizeWithoutFit)
                {
                    Console.WriteLine("Auto‑fitting rows increased the PDF file size.");
                }
                else if (sizeWithFit < sizeWithoutFit)
                {
                    Console.WriteLine("Auto‑fitting rows decreased the PDF file size.");
                }
                else
                {
                    Console.WriteLine("PDF file size unchanged after auto‑fitting rows.");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AutoFitRowsPdfComparison.Run();
        }
    }
}
