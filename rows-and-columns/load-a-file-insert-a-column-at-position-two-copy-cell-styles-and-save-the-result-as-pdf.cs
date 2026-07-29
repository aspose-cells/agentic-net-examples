// Title: C# – Insert a Column, Copy Formatting, and Export to PDF with Aspose.Cells
// Description: Loads an existing Excel file, inserts a new column at index 1 (second column), copies the style and data from column A to the new column, saves the workbook to a temporary XLSX, converts it to PDF using ConversionUtility.Convert, and removes the temporary file. Demonstrates column manipulation and PDF export in Aspose.Cells for .NET.
// Keywords: Aspose.Cells insert column C# | CopyColumn Aspose.Cells | Copy column formatting C# | ConversionUtility.Convert PDF | Excel to PDF Aspose.Cells | C# Excel column insertion | temporary file conversion Aspose | Aspose.Cells PDF export example
// Common Searches: how to insert a column at index 1 with Aspose.Cells | copy column style from column A using Aspose.Cells C# | convert modified workbook to PDF with Aspose.Cells | Aspose.Cells example insert column and export PDF | C# code for column insertion and PDF conversion
// Developer Intent: Add a new column, duplicate the formatting of an existing column, and generate a PDF from the updated workbook.
// Use Cases: Generating PDF reports where a placeholder column must match existing formatting before conversion. | Batch processing of Excel files that require a new column with identical style, followed by PDF output. | Preserving original workbooks by applying changes to a temporary file, converting to PDF, and cleaning up.
// AI Prompts: Write C# code that inserts a column at position two, copies the style from column A, and saves the workbook as a PDF using Aspose.Cells. | Explain the role of ConversionUtility.Convert when converting a temporary XLSX to PDF in Aspose.Cells. | Add comprehensive error handling for missing input files, column insertion failures, and PDF conversion errors in the described workflow.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an existing Excel file, inserts a new column at index 1 (second column), copies the style and data from column A to the new column, saves the workbook to a temporary XLSX, converts it to PDF using ConversionUtility.Convert, and removes the temporary file. Demonstrates column manipulation and PDF export in Aspose.Cells for .NET.
class InsertColumnAndSavePdf
{
    static void Main()
    {
        // Path to the existing Excel file
        string inputPath = "input.xlsx";

        // Load the workbook (create & load)
        Workbook workbook = new Workbook(inputPath);
        Worksheet sheet = workbook.Worksheets[0];
        Cells cells = sheet.Cells;

        // Insert a new column at position two (index 1, zero‑based)
        // Using the InsertColumn method rule
        cells.InsertColumn(1, true);

        // Copy styles (and data) from the original first column (index 0)
        // to the newly inserted column (index 1)
        // Using the CopyColumn method rule
        cells.CopyColumn(cells, 0, 1);

        // Save the modified workbook to a temporary Excel file
        string tempExcel = Path.Combine(Path.GetTempPath(), "temp_modified.xlsx");
        workbook.Save(tempExcel);

        // Convert the temporary Excel file to PDF using the ConversionUtility.Convert rule
        string outputPdf = "output.pdf";
        ConversionUtility.Convert(tempExcel, outputPdf);

        // Clean up the temporary file
        if (File.Exists(tempExcel))
        {
            File.Delete(tempExcel);
        }

        Console.WriteLine($"PDF saved to: {outputPdf}");
    }
}
