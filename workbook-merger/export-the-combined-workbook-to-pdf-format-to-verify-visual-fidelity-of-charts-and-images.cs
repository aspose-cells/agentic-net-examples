using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsPdfExport
{
    class Program
    {
        static void Main()
        {
            // Load the first workbook (creates a Workbook instance)
            Workbook combinedWorkbook = new Workbook("Workbook1.xlsx");

            // Load the second workbook to be merged
            Workbook secondWorkbook = new Workbook("Workbook2.xlsx");

            // Combine the second workbook into the first one
            combinedWorkbook.Combine(secondWorkbook);

            // Configure PDF save options (optional, e.g., export document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the combined workbook as a PDF file
            combinedWorkbook.Save("CombinedWorkbook.pdf", pdfOptions);

            Console.WriteLine("Combined workbook exported to PDF successfully.");
        }
    }
}