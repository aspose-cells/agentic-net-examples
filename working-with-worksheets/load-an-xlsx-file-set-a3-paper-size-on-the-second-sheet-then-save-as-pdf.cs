// Title: Set A3 Paper Size on the Second Worksheet and Export XLSX to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an XLSX workbook, verifies a second worksheet exists, changes that sheet's PageSetup.PaperSize to A3, and saves the entire workbook as a PDF file using Aspose.Cells.
// Keywords: Aspose.Cells | C# | .NET | set paper size | A3 | second worksheet | Excel to PDF | PageSetup | PDF export | Workbook.Save
// Common Searches: Aspose.Cells set A3 paper size on second sheet C# | Convert Excel workbook to PDF with custom page size per worksheet | How to change page setup of a specific worksheet before PDF export using Aspose.Cells | C# export XLSX to PDF with different paper sizes
// Developer Intent: Apply an A3 page size to the second worksheet of an Excel file and generate a PDF from the workbook.
// Use Cases: Produce printable reports where only the second sheet requires A3 formatting before PDF creation. | Automate batch conversion of Excel files to PDF while applying distinct page sizes to selected worksheets. | Create marketing brochures or large‑format documents from Excel where a particular sheet must be A3 for high‑resolution printing.
// AI Prompts: Generate C# code with Aspose.Cells to set the paper size of the third worksheet to Letter and export the workbook as PDF. | Explain how to assign different orientations, margins, and paper sizes to multiple worksheets before converting to PDF using Aspose.Cells. | Provide a step‑by‑step guide to validate worksheet existence, modify its PageSetup properties, and save the workbook as a PDF in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // Loads an XLSX workbook, verifies a second worksheet exists, changes that sheet's PageSetup.PaperSize to A3, and saves the entire workbook as a PDF file using Aspose.Cells.
    class SetPaperSizeAndConvertToPdf
    {
        static void Main()
        {
            // Path to the source XLSX file
            string sourceFile = "input.xlsx";

            // Path for the resulting PDF file
            string pdfFile = "output.pdf";

            // Load the workbook from the XLSX file
            Workbook workbook = new Workbook(sourceFile);

            // Ensure the workbook has at least two worksheets
            if (workbook.Worksheets.Count < 2)
                throw new InvalidOperationException("The workbook does not contain a second sheet.");

            // Access the second worksheet (index 1) and set its paper size to A3
            Worksheet secondSheet = workbook.Worksheets[1];
            secondSheet.PageSetup.PaperSize = PaperSizeType.PaperA3;

            // Save the modified workbook as a PDF document
            workbook.Save(pdfFile, SaveFormat.Pdf);

            Console.WriteLine($"Workbook saved as PDF with A3 paper size on the second sheet: {pdfFile}");
        }
    }
}
