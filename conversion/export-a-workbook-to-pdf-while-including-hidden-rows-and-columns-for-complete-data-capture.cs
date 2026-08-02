// Title: Export hidden rows, columns, and worksheets to PDF with Aspose.Cells for .NET
// Description: Demonstrates how to save an Aspose.Cells workbook as a PDF while preserving hidden rows, hidden columns, and hidden worksheets. The example uses PdfSaveOptions with SheetSet.All and adds a visible placeholder sheet to satisfy visibility requirements, ensuring all data appears in the generated PDF.
// Keywords: Aspose.Cells PDF export hidden rows | include hidden columns in PDF C# | PdfSaveOptions SheetSet.All | export hidden worksheets to PDF | Aspose.Cells workbook to PDF | .NET Excel to PDF conversion | C# Aspose.Cells hidden data
// Common Searches: Aspose.Cells export hidden rows to PDF C# | PdfSaveOptions include hidden columns | How to render hidden worksheets in PDF with Aspose.Cells | C# convert Excel with hidden data to PDF | Aspose.Cells PDF conversion hidden sheets
// Developer Intent: Generate a PDF from an Excel workbook that contains hidden rows, columns, or worksheets, ensuring the PDF captures all underlying data.
// Use Cases: Create a printable report that includes data hidden for UI layout purposes. | Archive an Excel file for compliance while preserving every cell in the PDF version. | Combine multiple hidden worksheets into a single PDF without exposing them in the source workbook.
// AI Prompts: Show how to configure PdfSaveOptions to keep cell formatting while exporting hidden rows and columns. | Provide a C# example that merges several hidden worksheets into one PDF using Aspose.Cells. | Explain how to export hidden rows and columns without adding a placeholder visible worksheet.

using System;
using Aspose.Cells;
using Aspose.Cells.Rendering; // Required for SheetSet

// Demonstrates how to save an Aspose.Cells workbook as a PDF while preserving hidden rows, hidden columns, and hidden worksheets. The example uses PdfSaveOptions with SheetSet.All and adds a visible placeholder sheet to satisfy visibility requirements, ensuring all data appears in the generated PDF.
class ExportPdfWithHiddenRowsAndColumns
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet dataSheet = workbook.Worksheets[0];
            dataSheet.Name = "DataSheet";

            // Populate some sample data
            dataSheet.Cells["A1"].PutValue("Header1");
            dataSheet.Cells["B1"].PutValue("Header2");
            dataSheet.Cells["C1"].PutValue("Header3");
            dataSheet.Cells["A2"].PutValue("R1C1");
            dataSheet.Cells["B2"].PutValue("R1C2");
            dataSheet.Cells["C2"].PutValue("R1C3");
            dataSheet.Cells["A3"].PutValue("R2C1");
            dataSheet.Cells["B3"].PutValue("R2C2");
            dataSheet.Cells["C3"].PutValue("R2C3");

            // Hide a row (row index 1 -> second row) and a column (column index 1 -> column B)
            dataSheet.Cells.HideRow(1);
            dataSheet.Cells.HideColumn(1);

            // Hide the data sheet but keep at least one visible worksheet in the workbook
            dataSheet.IsVisible = false;

            // Add an empty visible worksheet so the workbook satisfies the visibility requirement
            Worksheet visibleSheet = workbook.Worksheets[workbook.Worksheets.Add()];
            visibleSheet.Name = "VisiblePlaceholder";

            // Configure PDF save options to include all sheets (visible + hidden)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                SheetSet = SheetSet.All // Ensures hidden worksheets are also rendered
            };

            // Save the workbook as PDF; hidden rows and columns are retained in the output
            workbook.Save("Workbook_With_Hidden_Rows_Columns.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}
