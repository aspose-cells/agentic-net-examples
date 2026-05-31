using System;
using Aspose.Cells;

class InsertRowsAndExportPdf
{
    static void Main()
    {
        // Path to the existing workbook
        string inputFile = "input.xlsx";

        // Load the workbook (uses the Workbook(string) constructor)
        Workbook workbook = new Workbook(inputFile);

        // Prepare insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove, // copy format
            UpdateReference = true                       // update formulas/references
        };

        // Insert 5 rows starting at row index 30 (31st row, zero‑based index)
        // This uses the InsertRows(int, int, InsertOptions) overload
        workbook.Worksheets[0].Cells.InsertRows(30, 5, insertOptions);

        // Save the modified workbook as PDF (uses Workbook.Save(string, SaveFormat))
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}