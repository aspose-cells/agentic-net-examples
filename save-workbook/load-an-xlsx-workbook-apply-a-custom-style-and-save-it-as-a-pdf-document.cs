using System;
using System.Drawing;
using Aspose.Cells;

class Program
{
    static void Main()
    {
        // Load the existing XLSX workbook
        string sourcePath = "input.xlsx";
        Workbook workbook = new Workbook(sourcePath);

        // Create a custom style
        Style customStyle = workbook.CreateStyle();
        customStyle.Font.Name = "Arial";
        customStyle.Font.Size = 12;
        customStyle.ForegroundColor = Color.LightYellow;
        customStyle.Pattern = BackgroundType.Solid;

        // Apply the custom style to a range (A1:B10) on the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];
        Aspose.Cells.Range cellRange = worksheet.Cells.CreateRange("A1", "B10");
        StyleFlag flag = new StyleFlag { All = true };
        cellRange.ApplyStyle(customStyle, flag);

        // Configure PDF save options (optional: export document structure)
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            ExportDocumentStructure = true
        };

        // Save the workbook as a PDF file using the Save method with options
        workbook.Save("output.pdf", pdfOptions);
    }
}