using System;
using Aspose.Cells;
using Aspose.Cells.Saving;

class WorkbookToPdfWithoutComments
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Remove all comments from every worksheet
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            sheet.ClearComments();
        }

        // Optionally remove other personal information (author names, etc.)
        workbook.RemovePersonalInformation();

        // Configure PDF save options (default settings are sufficient for this task)
        PdfSaveOptions pdfOptions = new PdfSaveOptions();

        // Save the cleaned workbook as a PDF file
        workbook.Save("output.pdf", pdfOptions);

        Console.WriteLine("Workbook has been saved to PDF without comments.");
    }
}