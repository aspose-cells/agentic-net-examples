// Title: Merge Cells P15:Q17, Apply Date Format (mm‑dd‑yyyy) and Export to PDF with Aspose.Cells for .NET
// Description: C# example that creates a workbook, merges the range P15:Q17, sets the upper‑left cell to the custom date format "mm-dd-yyyy", saves the sheet as a temporary XLSX file, converts it to PDF using Aspose.Cells ConversionUtility, and cleans up the temporary file.
// Keywords: Aspose.Cells merge cells | C# merge P15 Q17 | custom date format mm-dd-yyyy | Aspose.Cells export to PDF | temporary XLSX to PDF conversion | .NET spreadsheet PDF generation | cell style custom format Aspose
// Common Searches: how to merge cells and set date format with Aspose.Cells | Aspose.Cells convert merged range to PDF | C# set custom number format for merged cells | Aspose.Cells PDF export after cell styling | merge P15 Q17 Aspose.Cells example
// Developer Intent: Create a workbook, merge a specific range, apply a date format, and generate a PDF file programmatically.
// Use Cases: Produce PDF invoices where the issue date spans two columns. | Generate calendar PDFs with month titles merged and formatted as dates. | Automate reporting templates that require a merged header cell displaying a formatted date.
// AI Prompts: Write C# code using Aspose.Cells to merge cells P15:Q17, set the number format to mm-dd-yyyy, and save the result as a PDF. | Explain the steps for converting a temporary XLSX workbook to PDF after applying a custom date style with Aspose.Cells. | Provide error‑handling best practices for temporary file cleanup when exporting a formatted workbook to PDF in .NET.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;

// C# example that creates a workbook, merges the range P15:Q17, sets the upper‑left cell to the custom date format "mm-dd-yyyy", saves the sheet as a temporary XLSX file, converts it to PDF using Aspose.Cells ConversionUtility, and cleans up the temporary file.
class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells P15:Q17
        // P column = index 15, row 15 = index 14, total rows = 3, total columns = 2
        worksheet.Cells.Merge(14, 15, 3, 2);

        // Set the number format of the merged cell (upper‑left cell P15) to "mm-dd-yyyy"
        Cell mergedCell = worksheet.Cells[14, 15];
        Style style = mergedCell.GetStyle();
        style.Custom = "mm-dd-yyyy";
        mergedCell.SetStyle(style);

        // Save the workbook to a temporary XLSX file (required for the conversion utility)
        string tempXlsxPath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".xlsx");
        workbook.Save(tempXlsxPath, SaveFormat.Xlsx);

        // Convert the temporary XLSX file to PDF using the provided ConversionUtility.Convert method
        string outputPdfPath = "MergedCellsOutput.pdf";
        ConversionUtility.Convert(tempXlsxPath, outputPdfPath);

        // Clean up the temporary XLSX file
        if (File.Exists(tempXlsxPath))
        {
            File.Delete(tempXlsxPath);
        }

        Console.WriteLine($"PDF file created at: {Path.GetFullPath(outputPdfPath)}");
    }
}
