// Title: Delete the 8th row from an Excel workbook and export it as PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file with Aspose.Cells, removes row index 7 from the first worksheet, and saves the result as a PDF. | Show how to use Aspose.Cells Cells.DeleteRow and Workbook.Save with SaveFormat.Pdf to delete a specific row and convert the workbook to PDF. | Generate a .NET example that deletes the eighth row in an Excel sheet and then creates a PDF version of the modified workbook.
// Common Searches: Aspose.Cells C# delete row 8 then save as PDF | How to remove a specific row from an Excel file and export to PDF using Aspose.Cells | C# example for deleting a row in a worksheet before PDF conversion with Aspose.Cells | Delete row index 7 in first sheet Aspose.Cells and generate PDF output | Remove eighth row from Excel and convert to PDF using .NET Aspose.Cells
// Tags: Cells.DeleteRow Aspose.Cells | Aspose.Cells SaveFormat.Pdf conversion | delete row before PDF export Aspose.Cells | C# Excel row removal Aspose.Cells | Excel to PDF after row deletion .NET

using System;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // // Loads 'input.xlsx', deletes the 8th row (zero‑based index 7) from the first worksheet, and saves the modified workbook as 'output.pdf' using Aspose.Cells for .NET.
    class DeleteRowAndConvertToPdf
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Load the workbook from the file (load rule)
            Workbook workbook = new Workbook(sourcePath);

            // Delete the 8th row (zero‑based index 7) in the first worksheet
            // DeleteRow method is part of the Cells class (feature rule)
            workbook.Worksheets[0].Cells.DeleteRow(7);

            // Save the modified workbook as PDF (save rule with SaveFormat.Pdf)
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Row 8 deleted and workbook saved as PDF to '{pdfPath}'.");
        }
    }
}
