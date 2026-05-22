using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Load the workbook (create rule)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove,
            UpdateReference = true
        };

        // Insert a single row at index 20 (21st row) with the specified options
        worksheet.Cells.InsertRows(20, 1, insertOptions);

        // Save the modified workbook as PDF (save rule; extension determines format)
        string outputPdf = "output.pdf";
        workbook.Save(outputPdf);
    }
}