// Title: C# – Insert Row at Index 20 with Formatting and Export Excel to PDF using Aspose.Cells
// Description: Loads an Excel workbook, adds a row at the 21st position while copying the style from the preceding row (CopyFormatType.SameAsAbove) and updating formulas, saves to a temporary file, converts it to PDF via ConversionUtility, and then removes the temporary file.
// Keywords: Aspose.Cells InsertRows | CopyFormatType.SameAsAbove | C# insert row Excel | Excel to PDF Aspose.Cells | update references after row insert | temporary file cleanup Aspose | ConversionUtility PDF | worksheet.Cells.InsertRows | preserve formatting Excel C# | Aspose.Cells .NET PDF conversion
// Common Searches: Aspose.Cells insert row at specific index C# | How to copy formatting when inserting rows with Aspose.Cells | Convert modified Excel workbook to PDF using Aspose.Cells | InsertRows with UpdateReference example | C# Aspose.Cells delete temporary file after PDF conversion
// Developer Intent: Add a row at position 21, retain the above row’s formatting and formulas, then generate a PDF from the updated workbook.
// Use Cases: Add a header row before exporting a financial report to PDF. | Insert a blank line for comments in a timesheet and produce a printable PDF. | Re‑structure a data sheet by adding summary rows and archive the result as PDF. | Automate report generation where layout changes require row insertion prior to PDF export.
// AI Prompts: Write C# code that inserts multiple rows at a given index, copies formatting from the preceding rows, updates all formulas, and saves the workbook directly to PDF with Aspose.Cells. | Show how to use InsertOptions with CopyFormatType.SameAsAbove and UpdateReference to maintain cell references after inserting rows. | Provide a robust pattern for converting an Aspose.Cells workbook to PDF while handling temporary file creation and deletion safely.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel workbook, adds a row at the 21st position while copying the style from the preceding row (CopyFormatType.SameAsAbove) and updating formulas, saves to a temporary file, converts it to PDF via ConversionUtility, and then removes the temporary file.
class InsertRowsAndConvertToPdf
{
    static void Main()
    {
        // Paths for the source Excel file and the final PDF output
        string sourceExcelPath = "input.xlsx";
        string intermediateExcelPath = "temp_modified.xlsx";
        string outputPdfPath = "output.pdf";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourceExcelPath);

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Set up insert options to copy formatting from the row above
        InsertOptions insertOptions = new InsertOptions
        {
            CopyFormatType = CopyFormatType.SameAsAbove, // copy format from the row above
            UpdateReference = true                       // update formulas/references
        };

        // Insert a single row at zero‑based index 20 (i.e., the 21st row)
        worksheet.Cells.InsertRows(20, 1, insertOptions);

        // Save the modified workbook to a temporary file
        workbook.Save(intermediateExcelPath);

        // Convert the temporary Excel file to PDF
        ConversionUtility.Convert(intermediateExcelPath, outputPdfPath);

        // Optional: clean up the temporary file
        try
        {
            System.IO.File.Delete(intermediateExcelPath);
        }
        catch
        {
            // Ignore any errors during cleanup
        }

        Console.WriteLine("Rows inserted and PDF generated successfully.");
    }
}
