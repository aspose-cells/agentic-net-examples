using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsTxtToPdf
{
    class Program
    {
        static void Main()
        {
            // Path to the source TXT (CSV/TSV) file
            string txtFilePath = "input.txt";

            // Load the TXT file into a workbook
            // Aspose.Cells automatically detects the format (CSV/TSV) based on the file extension
            Workbook workbook = new Workbook(txtFilePath);

            // Get the first worksheet (the TXT data is loaded into the first sheet)
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the last row that contains data
            int lastDataRow = sheet.Cells.MaxDataRow; // zero‑based index

            // Insert a horizontal page break after each row (paragraph) for clarity
            // Adding a break at the first cell of the next row forces the current row onto its own page
            for (int row = 0; row <= lastDataRow; row++)
            {
                // Build the cell name for the start of the next row (e.g., "A2", "A3", ...)
                string cellName = $"A{row + 2}"; // +2 because rows are 0‑based and we want the next row
                sheet.AddPageBreaks(cellName);
            }

            // Configure PDF save options (optional customizations can be added here)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: ensure each sheet is not forced onto a single page
                OnePagePerSheet = false
            };

            // Save the modified workbook as a PDF
            string pdfOutputPath = "output.pdf";
            workbook.Save(pdfOutputPath, pdfOptions);

            Console.WriteLine($"TXT workbook converted to PDF with page breaks: {pdfOutputPath}");
        }
    }
}