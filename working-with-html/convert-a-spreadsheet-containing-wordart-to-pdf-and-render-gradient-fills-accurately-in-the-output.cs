// Title: Convert WordArt with Gradient Fill to PDF using Aspose.Cells for .NET
// Description: Creates an Excel workbook, inserts a WordArt shape with a custom two‑color diagonal gradient, forces the fill type to Gradient, saves the file, and converts it to PDF while preserving the gradient appearance.
// Keywords: Aspose.Cells | WordArt gradient | Excel to PDF conversion | preserve gradient fill | C# Aspose.Cells example | ConversionUtility PDF | gradient WordArt export
// Common Searches: Aspose.Cells export WordArt gradient to PDF | C# preserve WordArt fill when converting Excel to PDF | how to add gradient WordArt in Aspose.Cells | convert Excel workbook with WordArt to PDF | gradient fill not showing in PDF Aspose.Cells
// Developer Intent: Generate a PDF from an Excel file that contains a WordArt object with a custom gradient, ensuring the gradient renders correctly in the output.
// Use Cases: Design marketing flyers in Excel with gradient WordArt titles and deliver them as PDF brochures. | Automate financial reports where section headers use gradient WordArt that must remain unchanged after PDF export. | Batch‑process archived Excel documents that include WordArt, converting each to PDF while keeping visual fidelity.
// AI Prompts: Show how to apply a three‑color gradient to WordArt before PDF conversion. | Provide code to convert the workbook to PDF directly without creating a temporary XLSX file. | Explain how to modify the gradient angle or style for WordArt shapes during export.

using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Utility;

// Creates an Excel workbook, inserts a WordArt shape with a custom two‑color diagonal gradient, forces the fill type to Gradient, saves the file, and converts it to PDF while preserving the gradient appearance.
class WordArtToPdf
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Add a WordArt shape with a preset style that includes a gradient fill
        // Parameters: style, text, topRow, top (pixels), leftColumn, left (pixels), height (pixels), width (pixels)
        Shape wordArt = worksheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle6, // Gradient Fill - Gray (preset)
            "Gradient WordArt",
            2,   // topRow
            0,   // top offset in pixels
            2,   // leftColumn
            0,   // left offset in pixels
            100, // height in pixels
            400  // width in pixels
        );

        // Ensure the fill type is set to Gradient
        wordArt.Fill.FillType = FillType.Gradient;

        // Apply a custom two‑color gradient to the WordArt for accurate rendering
        // Gradient from Red to Blue, diagonal down style, variant 1
        wordArt.Fill.SetTwoColorGradient(
            Color.Red,
            Color.Blue,
            GradientStyleType.DiagonalDown,
            1
        );

        // Save the workbook to a temporary Excel file (required for the conversion utility)
        string tempExcelPath = "temp_wordart.xlsx";
        workbook.Save(tempExcelPath);

        // Convert the Excel file to PDF, preserving the gradient fill in the WordArt
        string pdfOutputPath = "WordArtWithGradient.pdf";
        ConversionUtility.Convert(tempExcelPath, pdfOutputPath);
    }
}
