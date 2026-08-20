// Title: C# – Merge Cells P15:Q17, Apply Date Format mm‑dd‑yyyy, and Export Workbook to PDF with Aspose.Cells
// Description: Loads an XLSX workbook, merges the range P15:Q17 on the first worksheet, sets the merged cell's custom date format to "mm-dd-yyyy", saves the workbook, converts it to PDF using Aspose.Cells, and cleans up temporary files. Demonstrates end‑to‑end formatting and PDF export in .NET.
// Keywords: Aspose.Cells C# merge cells | set custom date format Aspose.Cells | export XLSX to PDF .NET | merge range P15:Q17 Aspose | Aspose.Cells PDF conversion example | C# workbook formatting Aspose | date format mm-dd-yyyy Aspose.Cells
// Common Searches: how to merge cells and set date format with Aspose.Cells | Aspose.Cells example merge P15 Q17 and export PDF | C# set custom number format for merged cells Aspose | convert formatted Excel to PDF using Aspose.Cells .NET | Aspose.Cells merge range and apply date style
// Developer Intent: Merge a specific cell range, apply a custom date format, and generate a PDF from the workbook using Aspose.Cells for .NET.
// Use Cases: Create a report header that spans P15:Q17, shows a date in mm‑dd‑yyyy format, and deliver the report as a PDF. | Build an invoice template where the due‑date cell is merged and formatted, then export the final invoice to PDF automatically. | Generate a schedule worksheet with a merged date cell and provide a ready‑to‑share PDF version for stakeholders.
// AI Prompts: Give a concise Aspose.Cells for .NET snippet that merges P15:Q17, applies the "mm-dd-yyyy" date format, and saves directly to PDF without a temporary file. | Explain how to use the Style.Custom property to set a date format on a merged cell and then convert the workbook to PDF with Aspose.Cells. | Show best practices for cleaning up temporary files after converting an XLSX workbook with merged cells to PDF using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells;

// Loads an XLSX workbook, merges the range P15:Q17 on the first worksheet, sets the merged cell's custom date format to "mm-dd-yyyy", saves the workbook, converts it to PDF using Aspose.Cells, and cleans up temporary files. Demonstrates end‑to‑end formatting and PDF export in .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells P15:Q17
        // P = column index 15 (zero‑based), Q = 16
        // Row 15 = index 14, Row 17 = index 16, total rows = 3, total columns = 2
        worksheet.Cells.Merge(14, 15, 3, 2);

        // Set the number format of the merged cell (upper‑left cell) to date "mm-dd-yyyy"
        Cell mergedCell = worksheet.Cells[14, 15];
        Style style = mergedCell.GetStyle();
        style.Custom = "mm-dd-yyyy";
        mergedCell.SetStyle(style);

        // Save the modified workbook to a temporary XLSX file
        string tempXlsx = Path.GetTempFileName().Replace(".tmp", ".xlsx");
        workbook.Save(tempXlsx);

        // Convert the temporary XLSX file to PDF using the provided ConversionUtility rule
        string outputPdf = "output.pdf";
        ConversionUtility.Convert(tempXlsx, outputPdf);

        // Clean up the temporary file
        if (File.Exists(tempXlsx))
        {
            File.Delete(tempXlsx);
        }

        Console.WriteLine($"Workbook processed and saved as PDF: {outputPdf}");
    }
}
