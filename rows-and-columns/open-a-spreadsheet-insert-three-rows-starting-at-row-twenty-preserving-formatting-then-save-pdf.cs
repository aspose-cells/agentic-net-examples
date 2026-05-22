using System;
using Aspose.Cells;

class InsertRowsAndSavePdf
{
    static void Main()
    {
        // Input Excel file path
        string inputFile = "input.xlsx";

        // Output PDF file path
        string outputFile = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to preserve formatting (copy format from the row above)
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert three rows starting at row 20 (zero‑based index 19)
        worksheet.Cells.InsertRows(19, 3, insertOptions);

        // Save the modified workbook as PDF
        workbook.Save(outputFile, SaveFormat.Pdf);
    }
}