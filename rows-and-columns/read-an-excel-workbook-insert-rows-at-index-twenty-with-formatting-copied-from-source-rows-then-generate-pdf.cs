// Title: C# – Insert Rows with Formatting and Export Excel to PDF using Aspose.Cells
// Description: Load an existing workbook, insert five rows at zero‑based index 20 while copying the style from the row above and updating formulas, then save the result directly as a PDF.
// Keywords: Aspose.Cells InsertRows | CopyFormatType SameAsAbove | C# Excel to PDF conversion | update formula references Aspose.Cells | zero based row index | Aspose.Cells SaveFormat.Pdf
// Common Searches: Aspose.Cells insert multiple rows at a specific position | preserve row formatting when inserting rows in .NET | convert modified Excel workbook to PDF with Aspose.Cells | how to update formulas after inserting rows in Excel using C#
// Developer Intent: Add five rows at row 21, inherit the formatting of the preceding row, adjust any dependent formulas, and generate a PDF from the updated worksheet.
// Use Cases: Add placeholder rows in a report template before populating data while keeping the original style. | Insert rows in a financial model and automatically shift formula references to maintain calculations. | Produce a PDF snapshot of a worksheet after structural changes such as row insertion, ensuring visual consistency.
// AI Prompts: Generate C# code with Aspose.Cells to insert rows at index 20, copy the above row's formatting, update references, and save as PDF. | Explain the effect of InsertOptions.CopyFormatType.SameAsAbove on PDF rendering in Aspose.Cells. | Step‑by‑step guide: load an Excel file, insert rows with formatting preservation, and convert the workbook to PDF using Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Load an existing workbook, insert five rows at zero‑based index 20 while copying the style from the row above and updating formulas, then save the result directly as a PDF.
class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputFile = "input.xlsx";

        // Path for the resulting PDF file
        string outputPdf = "output.pdf";

        // Load the workbook (lifecycle rule: load)
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index as needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Prepare insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove, // copy format from the row above
            UpdateReference = true                       // update formulas/references if any
        };

        // Insert 5 rows starting at row index 20 (zero‑based) with the specified options
        // Adjust the second parameter to the number of rows you need to insert
        worksheet.Cells.InsertRows(20, 5, insertOptions);

        // Save the workbook directly as PDF (lifecycle rule: save)
        workbook.Save(outputPdf, SaveFormat.Pdf);
    }
}
