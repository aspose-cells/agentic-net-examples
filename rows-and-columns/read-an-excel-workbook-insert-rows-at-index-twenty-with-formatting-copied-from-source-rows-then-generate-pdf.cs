// Title: Insert multiple rows with inherited formatting at row 21 and export the worksheet to PDF using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads an .xlsx file, inserts five rows starting at row 21 while copying the style from the preceding row, updates any formula references, and saves the result as a PDF with Aspose.Cells. | Demonstrate how to use Aspose.Cells InsertOptions with CopyFormatType.SameAsAbove to add rows and then convert the workbook to PDF in a single operation. | Generate a snippet that shows inserting rows at a specific index, preserving cell formatting and references, and calling Workbook.Save with SaveFormat.Pdf.
// Common Searches: how to insert rows with same formatting as above using Aspose.Cells C# | Aspose.Cells C# insert multiple rows and keep formulas then save as PDF | convert Excel worksheet to PDF after adding rows with inherited style Aspose.Cells
// Tags: CopyFormatType.SameAsAbove row insertion Aspose.Cells | update cell references after row insertion Aspose.Cells | export worksheet to PDF Aspose.Cells | row insertion at specific index Aspose.Cells | batch insert rows Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// // Loads input.xlsx, inserts five rows at row 21 copying the style from the row above and updating formulas, then saves the worksheet directly as output.pdf in PDF format using Aspose.Cells.
class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the source Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputPdf = "output.pdf";

        // Load the workbook from the file (lifecycle rule: load)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Configure insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove, // copy format from the row above
            UpdateReference = true                       // update formulas/references if any
        };

        // Insert 5 rows starting at row index 20 (zero‑based) with the specified options
        worksheet.Cells.InsertRows(20, 5, insertOptions);

        // Save the modified workbook directly as PDF (lifecycle rule: save)
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
