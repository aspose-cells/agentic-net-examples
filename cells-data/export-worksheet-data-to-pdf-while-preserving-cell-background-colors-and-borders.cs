// Title: Export an Excel worksheet to PDF while preserving cell background colors and borders with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an .xlsx file, applies a solid light‑blue fill and thin black borders to a specific range, and saves the sheet as a PDF using Aspose.Cells while keeping the visual styling. | Demonstrate how to configure the PdfSaveOptions object in Aspose.Cells so that the generated PDF retains all cell fills, border lines, and layout. | Create a C# routine that loops through all worksheets in a workbook and merges them into one PDF document, ensuring that every cell’s formatting is maintained.
// Common Searches: Aspose.Cells C# export worksheet to PDF preserving cell colors and borders | How to keep Excel cell formatting when converting to PDF with Aspose.Cells .NET | PdfSaveOptions configuration to retain background fill and borders in PDF output | Export multiple Excel sheets to a single PDF while maintaining styling using Aspose.Cells | C# convert XLSX to PDF with Aspose.Cells keeping grid lines and cell borders
// Tags: Aspose.Cells PDF export preserving cell styles | background fill retention in PDF conversion | border rendering with PdfSaveOptions | C# combine worksheets into single PDF Aspose.Cells | grid line rendering option Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// The program loads input.xlsx, applies a light‑blue solid fill and thin black borders to range A1:C5, then saves the first worksheet as output.pdf using PdfSaveOptions that preserve the original layout, cell background colors, and borders.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input workbook exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file '{inputPath}' was not found.");
                return;
            }

            // Load the workbook containing the worksheet to export
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Example: ensure some cells have background colors and borders
            // (Remove or modify this block if the source file already contains formatting)
            Style style = workbook.CreateStyle();
            style.ForegroundColor = System.Drawing.Color.LightBlue;
            style.Pattern = BackgroundType.Solid;
            style.Borders[BorderType.BottomBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.BottomBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.LeftBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.LeftBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.RightBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.RightBorder].Color = System.Drawing.Color.Black;
            style.Borders[BorderType.TopBorder].LineStyle = CellBorderType.Thin;
            style.Borders[BorderType.TopBorder].Color = System.Drawing.Color.Black;

            // Apply the style to a sample range (A1:C5)
            Aspose.Cells.Range range = worksheet.Cells.CreateRange("A1:C5");
            range.ApplyStyle(style, new StyleFlag { All = true });

            // Configure PDF save options to preserve cell formatting
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                // Keep the original layout; do not force all columns onto one page
                AllColumnsInOnePagePerSheet = false
                // Background colors are preserved by default
                // Render grid lines if needed (optional)
                // RenderGridLines = true
            };

            // Export the worksheet to PDF while keeping background colors and borders
            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine($"PDF successfully saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
