// Title: Insert Row at Index 20 with Same‑Above Formatting and Export to PDF using Aspose.Cells for .NET
// Description: Loads an Excel workbook, inserts a single row at zero‑based index 20 on the first worksheet while copying the formatting of the row above and updating formula references, saves the change to a temporary file, and converts the result to PDF with Aspose.Cells ConversionUtility.
// Keywords: Aspose.Cells InsertRows C# | CopyFormatType.SameAsAbove | insert row at specific index Aspose.Cells | update formulas after row insertion | Excel to PDF conversion Aspose.Cells | .NET Excel row formatting | ConversionUtility Convert example
// Common Searches: How to insert a row at a specific index and keep formatting in Aspose.Cells .NET | Aspose.Cells InsertRows with CopyFormatType.SameAsAbove example | Convert modified Excel workbook to PDF after inserting rows | Update cell references when inserting rows using Aspose.Cells | C# code to add a row and export Excel to PDF with Aspose
// Developer Intent: Add a row at position 20, preserve the above row’s style and formulas, then generate a PDF from the updated workbook.
// Use Cases: Add a header row to a report template before creating a PDF for distribution. | Insert a blank data row in a financial sheet while keeping cell styles and formulas, then produce a PDF for stakeholder review. | Adjust an invoice layout by inserting rows with inherited formatting and export the final version to PDF for client delivery.
// AI Prompts: Write C# code that inserts multiple rows at a given index, copies formatting from the preceding rows, and converts the workbook to PDF using Aspose.Cells. | Explain how InsertOptions.CopyFormatType.SameAsAbove works and how to ensure formulas are updated after inserting rows in Aspose.Cells. | Provide a step‑by‑step tutorial for inserting rows, saving a temporary workbook, and converting it to PDF with ConversionUtility in Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel workbook, inserts a single row at zero‑based index 20 on the first worksheet while copying the formatting of the row above and updating formula references, saves the change to a temporary file, and converts the result to PDF with Aspose.Cells ConversionUtility.
class Program
{
    static void Main()
    {
        // Paths for the original Excel file, a temporary modified file, and the final PDF.
        string inputFile = "input.xlsx";
        string tempFile = "modified.xlsx";
        string pdfFile = "output.pdf";

        // Load the existing workbook (lifecycle rule: use Workbook(string) constructor).
        Workbook workbook = new Workbook(inputFile);

        // Prepare insert options to copy the formatting from the row above the insertion point.
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove, // copy formatting from the row above
            UpdateReference = true                       // update any formulas/references
        };

        // Insert a single row at index 20 (zero‑based) with the specified options.
        // This uses the InsertRows(int, int, InsertOptions) method.
        workbook.Worksheets[0].Cells.InsertRows(20, 1, insertOptions);

        // Save the modified workbook to a temporary file (lifecycle rule: use Workbook.Save(string)).
        workbook.Save(tempFile);

        // Convert the modified Excel file to PDF (utility rule: ConversionUtility.Convert(string, string)).
        ConversionUtility.Convert(tempFile, pdfFile);
    }
}
