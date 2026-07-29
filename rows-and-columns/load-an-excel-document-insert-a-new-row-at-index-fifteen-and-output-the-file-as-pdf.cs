// Title: C# – Insert Row at Index 15 and Export Excel to PDF with Aspose.Cells
// Description: Load an Excel workbook, insert a blank row at zero‑based index 15 in the first worksheet, save the change, and convert the updated file to PDF using Aspose.Cells ConversionUtility.
// Keywords: Aspose.Cells insert row C# | Excel to PDF conversion Aspose | C# add row worksheet | ConversionUtility PDF Aspose | zero based row index Aspose.Cells
// Common Searches: Aspose.Cells insert row at specific index | C# convert modified Excel to PDF with Aspose | How to add a row before exporting to PDF using Aspose.Cells | Insert row 15 in Excel and save as PDF C#
// Developer Intent: Add a new row at position 15 in an Excel sheet and generate a PDF of the modified workbook.
// Use Cases: Insert a spacer row before a header, then create a PDF report. | Adjust pagination by adding rows before PDF export for consistent page breaks. | Automate invoice generation: add a line‑item row to a template and output a PDF.
// AI Prompts: Show C# code that inserts multiple rows at a given index and converts the workbook to PDF without using a temporary file. | Explain how to customize page settings in Aspose.Cells ConversionUtility after inserting rows.

using System;
using Aspose.Cells;
using Aspose.Cells.Utility;

// Load an Excel workbook, insert a blank row at zero‑based index 15 in the first worksheet, save the change, and convert the updated file to PDF using Aspose.Cells ConversionUtility.
class InsertRowAndConvertToPdf
{
    static void Main()
    {
        // Paths for the source Excel file, the intermediate modified file, and the final PDF.
        string sourceExcelPath = "input.xlsx";
        string modifiedExcelPath = "modified.xlsx";
        string outputPdfPath = "output.pdf";

        // Load the existing workbook.
        Workbook workbook = new Workbook(sourceExcelPath);

        // Insert a single row at index 15 (zero‑based) in the first worksheet.
        // This pushes the original row 15 down and creates a blank row at that position.
        Worksheet worksheet = workbook.Worksheets[0];
        worksheet.Cells.InsertRows(15, 1);

        // Save the workbook after modification to a temporary file.
        workbook.Save(modifiedExcelPath);

        // Convert the modified Excel file to PDF using the provided ConversionUtility rule.
        ConversionUtility.Convert(modifiedExcelPath, outputPdfPath);

        Console.WriteLine("Row inserted and PDF generated successfully.");
    }
}
