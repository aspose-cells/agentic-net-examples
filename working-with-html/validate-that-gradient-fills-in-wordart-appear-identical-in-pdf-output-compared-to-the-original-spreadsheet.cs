using System;
using System.Drawing;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Drawing.Texts;

class GradientWordArtPdfValidation
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // ------------------------------------------------------------
        // 1. Add a WordArt shape with a preset gradient style (WordArtStyle7)
        // ------------------------------------------------------------
        // Parameters: style, text, upperLeftRow, top, upperLeftColumn, left, height, width
        Shape wordArt = sheet.Shapes.AddWordArt(
            PresetWordArtStyle.WordArtStyle7,   // Gradient Fill - Blue, Accent 1, Reflection
            "Gradient WordArt",
            2, 0,                               // Row, top offset
            2, 0,                               // Column, left offset
            100, 400);                          // Height, width

        // ------------------------------------------------------------
        // 2. Verify that the shape is recognized as WordArt and has a gradient fill
        // ------------------------------------------------------------
        if (wordArt.IsWordArt)
        {
            // The TextEffect property gives access to text formatting (optional)
            TextEffectFormat textEffect = wordArt.TextEffect;
            textEffect.FontBold = true;
            textEffect.FontSize = 24;
        }

        // ------------------------------------------------------------
        // 3. Additionally, demonstrate setting a custom two‑color gradient
        // ------------------------------------------------------------
        // Add a regular rectangle shape to show explicit gradient configuration
        Shape rect = sheet.Shapes.AddRectangle(5, 0, 5, 0, 120, 200);
        rect.Fill.FillType = FillType.Gradient;                     // Enable gradient fill
        // Use FillFormat to set a two‑color gradient (LightSkyBlue → DarkBlue, horizontal)
        rect.Fill.SetTwoColorGradient(
            Color.LightSkyBlue,
            Color.DarkBlue,
            GradientStyleType.Horizontal,
            1);                                                       // Variant 1

        // ------------------------------------------------------------
        // 4. Save the workbook as Excel (XLSX) – original source
        // ------------------------------------------------------------
        string excelPath = "GradientWordArtDemo.xlsx";
        workbook.Save(excelPath, SaveFormat.Xlsx);

        // ------------------------------------------------------------
        // 5. Convert the same workbook to PDF
        // ------------------------------------------------------------
        string pdfPath = "GradientWordArtDemo.pdf";
        workbook.Save(pdfPath, SaveFormat.Pdf);

        // ------------------------------------------------------------
        // 6. Simple validation output
        // ------------------------------------------------------------
        Console.WriteLine("Workbook saved as XLSX: " + excelPath);
        Console.WriteLine("Workbook converted to PDF: " + pdfPath);
        Console.WriteLine("Please open both files and visually verify that the WordArt gradient appears identical.");
    }
}