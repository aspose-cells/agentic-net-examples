// Title: C# – Unhide Row 12, Set Height to 20 pt, and Export Workbook to PDF with Aspose.Cells
// Description: Load an Excel file, unhide the 12th row (index 11), set its height to 20 points, and save the workbook as a PDF using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# unhide row | set row height Aspose.Cells | export Excel to PDF .NET | unhide specific row PDF conversion | Aspose.Cells row visibility
// Common Searches: Aspose.Cells how to unhide a row and set height in C# | C# export Excel to PDF after changing row visibility | unhide row 12 Aspose.Cells and save as PDF
// Developer Intent: Unhide the 12th row, adjust its height to 20 points, and generate a PDF from the modified workbook.
// Use Cases: Reveal a hidden header before creating a client‑ready PDF report. | Adjust summary row height for better layout in exported PDFs. | Prepare a template by programmatically unhiding rows and archiving it as PDF.
// AI Prompts: Create a reusable C# method that accepts worksheet index, row number, height, and PDF path, then unhides the row, sets its height, and saves the workbook as PDF with Aspose.Cells. | Explain how to unhide multiple rows with distinct heights and export the result to PDF using Aspose.Cells for .NET. | Provide example code to unhide row 5, set its height to 15 points, and convert the workbook to PDF in C#.

using System;
using Aspose.Cells;

namespace AsposeCellsUnhideRowAndPdf
{
    // Load an Excel file, unhide the 12th row (index 11), set its height to 20 points, and save the workbook as a PDF using Aspose.Cells for .NET.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            // Replace "input.xlsx" with the actual path to your Excel file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];

            // Unhide row 12 (zero‑based index 11) and set its height to 20 points
            worksheet.Cells.UnhideRow(11, 20);

            // Save the modified workbook as a PDF document
            // The PDF will reflect the unhidden row with the specified height
            workbook.Save("output.pdf", SaveFormat.Pdf);
        }
    }
}
