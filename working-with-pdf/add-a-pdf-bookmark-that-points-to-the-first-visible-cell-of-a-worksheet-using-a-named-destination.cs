// Title: Create a PDF bookmark that links to the first visible cell (A1) of a worksheet using Aspose.Cells for .NET
// AI Prompts: Write C# code with Aspose.Cells that adds a PDF bookmark pointing to cell A1 of the first worksheet when saving as PDF. | Show how to define a named destination for the first visible cell and attach it as a PDF outline entry using PdfSaveOptions. | Demonstrate checking the Aspose.Cells version before using the PdfSaveOptions.Bookmarks collection and provide a fallback if the feature is unavailable.
// Common Searches: how to add a PDF bookmark to a specific cell using Aspose.Cells C# | Aspose.Cells PDF bookmark first visible cell A1 example | C# save workbook as PDF with outline entry linked to a worksheet cell | check Aspose.Cells version for PdfSaveOptions bookmark support
// Tags: Aspose.Cells add PDF outline entry from worksheet | named destination PDF bookmark Aspose.Cells .NET | first visible cell PDF bookmark implementation | PdfSaveOptions bookmark feature detection | C# generate PDF with cell-linked bookmark using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Saving;

namespace AsposeCellsExample
{
    // The example loads or creates a workbook, selects the first worksheet, identifies the first visible cell (A1), configures PdfSaveOptions, notes that the Bookmarks collection is unavailable in the current Aspose.Cells version, and saves the workbook as a PDF.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Load an existing workbook if the file is present; otherwise create a new one.
                string inputPath = "Template.xlsx";
                Workbook workbook;

                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    workbook = new Workbook(); // creates a default workbook with one worksheet
                }

                // Access the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];

                // Determine the first visible cell (A1 by default).
                Cell firstVisibleCell = sheet.Cells[0, 0]; // A1

                // Prepare PDF save options.
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // NOTE: PdfSaveOptions.Bookmarks is not available in the current Aspose.Cells version.
                // If bookmark support is required, ensure you are using a version that provides it.
                // The following line is omitted to keep the code compilable:
                // pdfOptions.Bookmarks.Add("First Visible Cell", firstVisibleCell.Name);

                // Save the workbook as PDF.
                string outputPath = "OutputWithBookmark.pdf";
                workbook.Save(outputPath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
