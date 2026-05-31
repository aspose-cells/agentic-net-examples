using System;
using Aspose.Cells;

class RemoveDuplicatesAndExportPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputFile = "output.pdf";

        // Load the workbook from the specified file
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Remove duplicate rows in the worksheet while keeping formulas intact
        worksheet.Cells.RemoveDuplicates();

        // Save the cleaned workbook as a PDF document
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}