using System;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the workbook from a file
        Workbook wb = new Workbook("input.xlsx");

        // Name of the worksheet that should be excluded from the PDF
        string sheetNameToRemove = "SecretSheet";

        // Find the index of the worksheet with the specified name
        int sheetIndex = -1;
        for (int i = 0; i < wb.Worksheets.Count; i++)
        {
            if (wb.Worksheets[i].Name.Equals(sheetNameToRemove, StringComparison.OrdinalIgnoreCase))
            {
                sheetIndex = i;
                break;
            }
        }

        // If the worksheet exists, remove it from the collection
        if (sheetIndex >= 0)
        {
            wb.Worksheets.RemoveAt(sheetIndex);
        }

        // Configure PDF save options (optional: ignore completely blank pages)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            PrintingPageType = PrintingPageType.IgnoreBlank
        };

        // Save the workbook as PDF; the removed sheet will not appear in the output
        wb.Save("output.pdf", pdfOptions);
    }
}

// Author: Aspose.Cells .NET example – removes a specific worksheet before PDF export.