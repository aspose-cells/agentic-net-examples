// Title: C# – Insert Row at Index 15 and Export Excel to PDF with Aspose.Cells
// Description: Loads an existing workbook, inserts a new row at zero‑based index 15 on the first worksheet, saves the change to a temporary file, and converts the result to PDF using Aspose.Cells ConversionUtility.
// Keywords: Aspose.Cells InsertRow C# | Excel to PDF conversion Aspose | Worksheet.Cells.InsertRow example | ConversionUtility Convert PDF | temporary workbook save | C# Excel row insertion
// Common Searches: Aspose.Cells insert row at specific index | C# convert modified Excel to PDF | How to add a row before exporting to PDF with Aspose | Insert row 15 in Excel using Aspose.Cells .NET | Save workbook temporarily then generate PDF
// Developer Intent: Add a row at a fixed position in an Excel sheet and produce a PDF of the updated file.
// Use Cases: Add a header row to a financial statement before creating a PDF report. | Insert a blank spacer row in a data table to improve layout for a PDF invoice. | Programmatically modify a template by inserting rows at a known index and then generate a PDF for automated distribution.
// AI Prompts: Generate C# code that opens an Excel file, inserts a row at index 15, saves to a temporary workbook, and converts it to PDF using Aspose.Cells ConversionUtility. | Explain how Worksheet.Cells.InsertRow works with zero‑based indexing and how to chain it with ConversionUtility.Convert for PDF output. | Provide a robust C# example that includes error handling, temporary file cleanup, and logging when inserting rows and exporting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an existing workbook, inserts a new row at zero‑based index 15 on the first worksheet, saves the change to a temporary file, and converts the result to PDF using Aspose.Cells ConversionUtility.
class InsertRowAndConvertToPdf
{
    static void Main()
    {
        // Paths for the original Excel file, a temporary modified file, and the final PDF.
        string inputFile = "input.xlsx";
        string tempFile = "temp_modified.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook.
        Workbook workbook = new Workbook(inputFile);

        // Access the first worksheet (you can change the index if needed).
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at index 15 (zero‑based). Existing rows from 15 onward are shifted down.
        sheet.Cells.InsertRow(15);

        // Save the modified workbook to a temporary file.
        workbook.Save(tempFile);

        // Convert the temporary Excel file to PDF using the provided ConversionUtility rule.
        ConversionUtility.Convert(tempFile, outputPdf);

        // Optional: clean up the temporary file.
        // System.IO.File.Delete(tempFile);

        Console.WriteLine("Row inserted and PDF generated successfully.");
    }
}
