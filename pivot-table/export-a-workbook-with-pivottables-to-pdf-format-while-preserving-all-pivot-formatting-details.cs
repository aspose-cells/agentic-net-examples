// Title: Export a Formatted PivotTable to PDF with Aspose.Cells for .NET (C#)
// Description: Creates a workbook, adds sample sales data, builds a pivot table, applies a custom style, enables PreserveFormatting, refreshes the pivot, sets PdfSaveOptions.ExportDocumentStructure, and saves the file as a PDF while keeping all pivot formatting intact.
// Keywords: Aspose.Cells | C# pivot table PDF export | preserve pivot formatting | PdfSaveOptions ExportDocumentStructure | styled pivot table PDF | Aspose.Cells PivotTable formatting | export Excel pivot to PDF | Aspose.Cells PDFSaveOptions | pivot table style retention | .NET Excel to PDF
// Common Searches: Aspose.Cells keep pivot table formatting when exporting to PDF | C# export pivot table to PDF with styles | PdfSaveOptions ExportDocumentStructure example | How to preserve pivot table colors in PDF using Aspose.Cells | Export Excel pivot table as PDF .NET
// Developer Intent: Generate a PDF from an Excel workbook that contains a pivot table while retaining all custom pivot formatting.
// Use Cases: Create printable sales dashboards where pivot table headers and data cells keep bold fonts and background colors in the PDF. | Automate monthly reporting that embeds styled pivot tables, ensuring totals and subtotals appear exactly as in the original Excel file. | Deliver client‑facing PDF reports from Excel data without losing any pivot table visual customizations.
// AI Prompts: Show C# code using Aspose.Cells to export a workbook with a styled pivot table to PDF, preserving all custom formats. | Explain how to configure PdfSaveOptions.ExportDocumentStructure to retain pivot table layout when saving as PDF. | Demonstrate applying conditional formatting to a pivot table and keeping it after PDF export with Aspose.Cells.

using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Pivot;
using Aspose.Cells.Saving;

namespace AsposeCellsPivotPdfExport
{
    // Creates a workbook, adds sample sales data, builds a pivot table, applies a custom style, enables PreserveFormatting, refreshes the pivot, sets PdfSaveOptions.ExportDocumentStructure, and saves the file as a PDF while keeping all pivot formatting intact.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the pivot table
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Region");
            sheet.Cells["C1"].PutValue("Sales");

            sheet.Cells["A2"].PutValue("Bike");
            sheet.Cells["B2"].PutValue("North");
            sheet.Cells["C2"].PutValue(1200);

            sheet.Cells["A3"].PutValue("Bike");
            sheet.Cells["B3"].PutValue("South");
            sheet.Cells["C3"].PutValue(800);

            sheet.Cells["A4"].PutValue("Car");
            sheet.Cells["B4"].PutValue("North");
            sheet.Cells["C4"].PutValue(1500);

            sheet.Cells["A5"].PutValue("Car");
            sheet.Cells["B5"].PutValue("South");
            sheet.Cells["C5"].PutValue(1300);

            // Add a pivot table based on the data range
            int pivotIndex = sheet.PivotTables.Add("A1:C5", "E3", "SalesPivot");
            PivotTable pivotTable = sheet.PivotTables[pivotIndex];

            // Configure pivot fields
            pivotTable.AddFieldToArea(PivotFieldType.Row, "Product");
            pivotTable.AddFieldToArea(PivotFieldType.Column, "Region");
            pivotTable.AddFieldToArea(PivotFieldType.Data, "Sales");

            // Preserve formatting when the pivot table is refreshed
            pivotTable.PreserveFormatting = true;

            // Create a style to apply to the data area of the pivot table
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Font.Name = "Calibri";
            dataStyle.Font.Size = 11;
            dataStyle.Font.IsBold = true;
            dataStyle.ForegroundColor = Color.LightYellow;
            dataStyle.Pattern = BackgroundType.Solid;

            // Apply the style to the data area using PivotFormats
            // This ensures the formatting is part of the pivot table definition
            pivotTable.PivotFormats.FormatArea(
                PivotFieldType.Data,          // Target area
                0,                            // Subtotal index (0 = none)
                PivotFieldSubtotalType.None, // No subtotal
                PivotTableSelectionType.DataAndLabel, // Apply to data and label cells
                false,                        // Do not apply to grand totals
                false,                        // Do not apply to subtotals
                dataStyle);                   // The style to apply

            // Refresh all pivot tables in the workbook to ensure data and formatting are up‑to‑date
            workbook.Worksheets.RefreshPivotTables();

            // Set PDF save options to retain document structure (helps preserve pivot formatting)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as a PDF file
            workbook.Save("PivotTableExport.pdf", pdfOptions);
        }
    }
}
