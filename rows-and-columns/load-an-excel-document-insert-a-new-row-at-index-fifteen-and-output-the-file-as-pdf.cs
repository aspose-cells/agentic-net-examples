// Title: Insert Row 15 in Excel and Export to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel file using Aspose.Cells, inserts a new row at index 14 (the 15th row) in the first worksheet, saves the workbook, and converts it to PDF via ConversionUtility.
// Keywords: Aspose.Cells C# insert row | Aspose.Cells convert Excel to PDF | C# add row Excel worksheet | Aspose.Cells ConversionUtility | Excel row insertion PDF export | Aspose.Cells .NET | Excel to PDF conversion .NET
// Common Searches: C# insert row at specific position using Aspose.Cells | Aspose.Cells convert modified workbook to PDF | How to add a row in Excel and export as PDF with .NET | Aspose.Cells insert row before header and save PDF
// Developer Intent: Add a row at the 15th position in an Excel sheet and generate a PDF of the updated workbook.
// Use Cases: Insert a header row above existing data and produce a PDF report. | Add a blank spacer row in a financial statement before exporting to PDF. | Automate row insertion in an invoice template and create a PDF version for client distribution.
// AI Prompts: Show C# code that inserts a row at index 14 in an Aspose.Cells worksheet and saves the file. | Provide an example of converting a modified Excel workbook to PDF using Aspose.Cells.Utility.ConversionUtility. | Explain error‑handling strategies when inserting rows and converting to PDF with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an Excel file using Aspose.Cells, inserts a new row at index 14 (the 15th row) in the first worksheet, saves the workbook, and converts it to PDF via ConversionUtility.
class InsertRowAndConvertToPdf
{
    static void Main()
    {
        // Paths for the source Excel file, the intermediate modified file, and the final PDF.
        string sourceExcel = "input.xlsx";
        string modifiedExcel = "modified.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook.
        Workbook workbook = new Workbook(sourceExcel);

        // Access the first worksheet (you can change the index if needed).
        Worksheet sheet = workbook.Worksheets[0];

        // Insert a new row at the 15th position (zero‑based index 14).
        sheet.Cells.InsertRow(14);

        // Save the workbook after insertion.
        workbook.Save(modifiedExcel);

        // Convert the modified Excel file to PDF using the provided utility method.
        ConversionUtility.Convert(modifiedExcel, outputPdf);

        Console.WriteLine("Row inserted and PDF generated successfully.");
    }
}
