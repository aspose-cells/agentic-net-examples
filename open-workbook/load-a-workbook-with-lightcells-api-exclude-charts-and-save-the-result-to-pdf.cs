// Title: Load an Excel workbook with LightCells, remove all charts, and export to PDF using Aspose.Cells for .NET
// Description: Demonstrates how to open an .xlsx file with Aspose.Cells LoadOptions (compatible with LightCells), clear every chart from each worksheet, and save the modified workbook as a PDF document in C#.
// Keywords: Aspose.Cells LightCells | C# load Excel workbook | remove charts Excel | export Excel to PDF | chart‑free PDF conversion | Aspose.Cells LoadOptions | clear worksheet charts | PDF generation without charts
// Common Searches: Aspose.Cells LightCells load workbook and save as PDF | C# remove all charts from Excel before PDF export | How to exclude charts when converting Excel to PDF with Aspose.Cells | Clear worksheet charts using Aspose.Cells .NET | LightCells API PDF conversion without charts
// Developer Intent: Open an Excel file with LightCells, delete every chart object, and generate a PDF version of the workbook.
// Use Cases: Produce clean PDF reports from Excel templates that contain charts you don't want to display. | Batch‑convert multiple workbooks to PDF while stripping charts to reduce file size and speed up rendering. | Create PDF invoices or statements from Excel sheets where embedded charts should be omitted for a professional layout.
// AI Prompts: Write C# code that uses Aspose.Cells LoadOptions for LightCells to open an .xlsx file, clear all charts from each worksheet, and save the workbook as a PDF. | Show how to configure LightCells in Aspose.Cells and exclude charts during Excel‑to‑PDF conversion in .NET.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering;

namespace AsposeCellsLightCellsPdfExample
{
    // Demonstrates how to open an .xlsx file with Aspose.Cells LoadOptions (compatible with LightCells), clear every chart from each worksheet, and save the modified workbook as a PDF document in C#.
    class Program
    {
        static void Main()
        {
            // Path to the source Excel file
            string sourcePath = "input.xlsx";

            // Initialize LoadOptions (can be customized for LightCells if needed)
            LoadOptions loadOptions = new LoadOptions();

            // Load the workbook using the constructor that accepts a file path and LoadOptions
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Remove all charts from each worksheet to exclude them from the PDF output
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Clear the chart collection of the current worksheet
                sheet.Charts.Clear();
            }

            // Save the modified workbook as PDF using the provided Save method overload
            string pdfPath = "output.pdf";
            workbook.Save(pdfPath, SaveFormat.Pdf);

            Console.WriteLine($"Workbook loaded, charts excluded, and saved to PDF at: {pdfPath}");
        }
    }
}
