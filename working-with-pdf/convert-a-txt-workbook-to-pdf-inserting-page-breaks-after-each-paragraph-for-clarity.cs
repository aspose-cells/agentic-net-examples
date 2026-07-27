using System;
using System.IO;
using Aspose.Cells;

namespace TxtToPdfWithPageBreaks
{
    // Author: Aspose.Cells .NET example
    class Program
    {
        static void Main()
        {
            // Path to the source TXT file
            string txtPath = "source.txt";

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Read all lines from the TXT file
            string[] lines = File.ReadAllLines(txtPath);

            int currentRow = 0;
            bool previousLineWasEmpty = false;

            foreach (string line in lines)
            {
                // Write the line into column A (index 0)
                sheet.Cells[currentRow, 0].PutValue(line);
                currentRow++;

                // Detect paragraph boundaries (empty line)
                bool isEmpty = string.IsNullOrWhiteSpace(line);
                if (isEmpty && !previousLineWasEmpty && currentRow > 1)
                {
                    // Add a horizontal page break after the empty line
                    // The cell name for the break is the first cell of the next row (e.g., "A5")
                    string cellName = CellsHelper.CellIndexToName(currentRow, 0);
                    sheet.HorizontalPageBreaks.Add(cellName);
                }

                previousLineWasEmpty = isEmpty;
            }

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Example: ensure each sheet can span multiple pages
                OnePagePerSheet = false
            };
            workbook.Save("output.pdf", pdfOptions);
        }
    }
}