using System;
using Aspose.Cells;

public class ApplyZoomAndExportPdf
{
    public static void Main()
    {
        // Load the workbook (replace the path with your actual file)
        Workbook workbook = new Workbook("input.xlsx");

        // Iterate through each worksheet in the workbook
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // MaxDataRow is zero‑based; add 1 to get the total row count
            int rowCount = sheet.Cells.MaxDataRow + 1;

            // If the worksheet contains more than 500 rows, set zoom to 95%
            if (rowCount > 500)
            {
                sheet.Zoom = 95; // Worksheet.Zoom property
            }
        }

        // Create PDF save options (default settings)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);
    }
}