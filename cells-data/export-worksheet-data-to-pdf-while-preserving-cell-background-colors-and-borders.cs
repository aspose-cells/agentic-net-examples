// Title: Export an Excel worksheet to PDF while preserving cell background colors and borders using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that creates a workbook, styles a header row with a solid fill and thin borders, applies thin borders to data rows, and saves the worksheet as a PDF using Aspose.Cells PdfSaveOptions to retain the formatting. | Demonstrate how to set up Aspose.Cells PdfSaveOptions (e.g., ExportDocumentStructure) in a .NET application to keep cell styles such as background colors and borders during Excel‑to‑PDF conversion.
// Common Searches: C# Aspose.Cells export Excel to PDF preserving cell background color and borders | how to keep worksheet formatting when converting to PDF with Aspose.Cells .NET | Aspose.Cells PdfSaveOptions ExportDocumentStructure example for preserving styles | apply style to range and save as PDF using Aspose.Cells in C# | export formatted Excel sheet to PDF with Aspose.Cells without losing borders
// Tags: Aspose.Cells export worksheet to PDF with formatting | PdfSaveOptions ExportDocumentStructure preserve styles | apply solid fill and borders using StyleFlag Aspose.Cells | C# style range before PDF conversion Aspose.Cells | preserve cell background color PDF export .NET

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;

// The sample creates a workbook, fills it with data, applies a light‑blue background and thin borders to the header row and thin borders to the data rows, then saves the worksheet as a PDF using PdfSaveOptions with ExportDocumentStructure enabled, ensuring that cell colors and borders are retained in the generated PDF.
class ExportWorksheetToPdf
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Fill some sample data
            sheet.Cells["A1"].PutValue("Product");
            sheet.Cells["B1"].PutValue("Price");
            sheet.Cells["A2"].PutValue("Apple");
            sheet.Cells["B2"].PutValue(1.2);
            sheet.Cells["A3"].PutValue("Banana");
            sheet.Cells["B3"].PutValue(0.8);

            // ------------------------------
            // Apply background color and borders to the header row
            // ------------------------------
            Style headerStyle = workbook.CreateStyle();
            headerStyle.ForegroundColor = Color.LightBlue;          // background color
            headerStyle.Pattern = BackgroundType.Solid;            // solid fill
            headerStyle.Font.IsBold = true;                        // make text bold

            // Set thin borders on all sides
            headerStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            headerStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            // Apply the style to the header range A1:B1
            StyleFlag flag = new StyleFlag { All = true };
            sheet.Cells.CreateRange("A1:B1").ApplyStyle(headerStyle, flag);

            // ------------------------------
            // Apply borders to the data rows (optional background)
            // ------------------------------
            Style dataStyle = workbook.CreateStyle();
            dataStyle.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            dataStyle.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;

            sheet.Cells.CreateRange("A2:B3").ApplyStyle(dataStyle, flag);

            // ------------------------------
            // Configure PDF save options to retain document structure
            // (background colors and borders are preserved by default)
            // ------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Define output file path
            string outputPath = "WorksheetWithFormatting.pdf";

            // Save the worksheet as a PDF file
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
