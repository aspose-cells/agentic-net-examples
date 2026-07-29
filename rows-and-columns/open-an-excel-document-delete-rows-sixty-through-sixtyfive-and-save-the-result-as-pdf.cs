// Title: C# Aspose.Cells: Remove rows 60‑65 from Excel and export to PDF
// Description: Loads an XLSX file with Aspose.Cells, deletes rows 60‑65 on the first worksheet (zero‑based index 59, count 6), saves the change to a temporary workbook, converts it to PDF via ConversionUtility, then removes the temporary file.
// Keywords: Aspose.Cells | C# Excel row deletion | delete rows 60 to 65 | Excel to PDF conversion | ConversionUtility | temporary workbook | save as PDF .NET | remove specific rows | Aspose.Cells SaveFormat.Pdf | worksheet row removal
// Common Searches: Aspose.Cells C# remove rows 60 through 65 | Convert edited Excel workbook to PDF with Aspose | How to delete a specific row range before PDF export in .NET | C# example for Excel row range removal and PDF generation | Aspose.Cells temporary file handling for PDF conversion
// Developer Intent: Strip rows 60‑65 from the first sheet of an Excel workbook and produce a PDF of the cleaned document.
// Use Cases: Generating client‑ready PDFs after eliminating placeholder rows | Automating data‑cleanup for regulatory reports before archiving | Batch processing multiple workbooks to remove header/footer rows and create PDF versions
// AI Prompts: Write a concise C# snippet using Aspose.Cells that deletes rows 60‑65 and saves directly to PDF without a temporary file. | Explain error‑handling strategies for the temporary file and conversion steps in the provided code. | Suggest performance optimizations for processing large workbooks when removing rows and converting to PDF.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Loads an XLSX file with Aspose.Cells, deletes rows 60‑65 on the first worksheet (zero‑based index 59, count 6), saves the change to a temporary workbook, converts it to PDF via ConversionUtility, then removes the temporary file.
class Program
{
    static void Main()
    {
        // Paths for the original Excel file, a temporary modified file, and the final PDF.
        string sourceExcel = "input.xlsx";
        string tempExcel = "temp_modified.xlsx";
        string outputPdf = "output.pdf";

        // Load the existing workbook (load rule).
        Workbook workbook = new Workbook(sourceExcel);

        // Access the first worksheet.
        Worksheet sheet = workbook.Worksheets[0];

        // Delete rows 60 through 65.
        // API uses zero‑based indices, so row 60 = index 59.
        // Total rows to delete = 6 (60,61,62,63,64,65).
        sheet.Cells.DeleteRows(59, 6);

        // Save the modified workbook to a temporary file (save rule).
        workbook.Save(tempExcel, SaveFormat.Xlsx);

        // Convert the temporary Excel file to PDF using the provided ConversionUtility (conversion rule).
        ConversionUtility.Convert(tempExcel, outputPdf);

        // Clean up the temporary file.
        if (File.Exists(tempExcel))
        {
            File.Delete(tempExcel);
        }

        Console.WriteLine("Rows 60‑65 deleted and PDF saved to " + outputPdf);
    }
}
