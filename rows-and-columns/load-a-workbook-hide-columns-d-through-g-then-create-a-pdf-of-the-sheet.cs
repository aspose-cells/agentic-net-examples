// Title: C# – Hide Columns D‑G in an Excel Worksheet and Export to PDF using Aspose.Cells
// Description: Load an existing workbook with Aspose.Cells for .NET, hide columns D through G on the first worksheet, and save the result directly as a PDF file. The example includes file‑existence checking and basic error handling.
// Keywords: Aspose.Cells hide columns C# | Excel to PDF conversion .NET | Hide columns D G Aspose | Cells.HideColumns example | Export hidden‑column sheet as PDF | Aspose.Cells PDF save format
// Common Searches: how to hide multiple columns in Excel with Aspose.Cells | Aspose.Cells C# hide columns D to G before PDF export | convert Excel to PDF while hiding specific columns | C# code to hide columns and save workbook as PDF | Aspose.Cells hide columns then save as PDF
// Developer Intent: Hide a range of columns in an Excel file and generate a PDF of the modified sheet using Aspose.Cells for .NET.
// Use Cases: Create client‑ready PDFs that omit confidential or calculation columns. | Automate batch processing of Excel templates where layout columns should not appear in the final PDF. | Produce printable reports from spreadsheets while removing helper columns that are only needed for internal use.
// AI Prompts: Generate C# code with Aspose.Cells to hide columns 3‑6 and export the worksheet to PDF. | Add comprehensive error handling for missing input files and save failures when converting Excel to PDF with hidden columns. | Show how to hide non‑contiguous column ranges before saving an Excel sheet as a PDF using Aspose.Cells for .NET.

using System;
using System.IO;
using Aspose.Cells;

// Load an existing workbook with Aspose.Cells for .NET, hide columns D through G on the first worksheet, and save the result directly as a PDF file. The example includes file‑existence checking and basic error handling.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found.");
                return;
            }

            // Load the workbook from the file
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet and its cells
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Hide columns D (index 3) through G (index 6) – total 4 columns
            cells.HideColumns(3, 4);

            // Save the modified workbook as a PDF
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
