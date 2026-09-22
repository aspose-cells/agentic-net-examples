// Title: How to set the PDF title metadata from an Excel workbook name using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file with Aspose.Cells, extracts the file name without its extension, assigns that value to the workbook's title property, and saves the workbook as a PDF. | Demonstrate how to configure PdfSaveOptions and set the built‑in document title before exporting a Workbook to PDF so the resulting file inherits the original Excel filename as its title.
// Common Searches: Aspose.Cells C# set PDF document title to Excel file name during conversion | How to copy workbook filename into PDF Title metadata with Aspose.Cells | Saving Excel as PDF with custom Title property using Aspose.Cells .NET | Set built‑in document property title before exporting workbook to PDF in C#
// Tags: Aspose.Cells set PDF document title | C# set workbook title property | Excel filename to PDF metadata conversion | PdfSaveOptions custom document properties | Extract workbook name for PDF title

using System;
using System.IO;
using Aspose.Cells;

namespace Example
{
    // The example checks for the presence of input.xlsx, loads it into an Aspose.Cells Workbook, derives the file name without extension, assigns this string to the workbook's built‑in Title property, and then saves the workbook as output.pdf using PdfSaveOptions, resulting in a PDF whose Title metadata matches the original Excel file name.
    class Program
    {
        static void Main()
        {
            string excelFilePath = "input.xlsx";

            // Verify that the input Excel file exists
            if (!File.Exists(excelFilePath))
            {
                Console.WriteLine($"Error: Excel file not found at path '{excelFilePath}'.");
                return;
            }

            try
            {
                // Load the original Excel workbook
                Workbook workbook = new Workbook(excelFilePath);

                // Derive the title from the workbook file name (without extension)
                string pdfTitle = Path.GetFileNameWithoutExtension(excelFilePath);

                // Set PDF title metadata via built‑in document properties
                workbook.BuiltInDocumentProperties.Title = pdfTitle;

                // Configure PDF save options
                PdfSaveOptions pdfOptions = new PdfSaveOptions();

                // Save the workbook as a PDF with the specified options
                string pdfFilePath = "output.pdf";
                workbook.Save(pdfFilePath, pdfOptions);

                Console.WriteLine($"PDF saved successfully to '{pdfFilePath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
