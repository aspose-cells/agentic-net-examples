// Title: Convert an Excel workbook to PDF and add a diagonal semi‑transparent CONFIDENTIAL WordArt watermark to every worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, inserts a 45‑degree rotated, 50% transparent WordArt shape with custom text on each worksheet, and saves the workbook as a PDF using Aspose.Cells. | Create a reusable C# method that accepts watermark text, font size, opacity, and rotation angle, applies the watermark to all worksheets in a workbook, and then exports the result to PDF with Aspose.Cells. | Show how to set the Z‑order of a WordArt shape so the watermark appears behind the cells before converting the workbook to PDF with Aspose.Cells.
// Common Searches: Aspose.Cells C# add diagonal text watermark to every worksheet before PDF conversion | How to set WordArt shape transparency and rotation in Aspose.Cells .NET | Export Excel workbook to PDF with semi‑transparent watermark using Aspose.Cells | Place watermark behind cells in PDF output with Aspose.Cells C# | Apply same watermark to multiple worksheets in Aspose.Cells example
// Tags: add watermark to worksheets Aspose.Cells | export workbook to PDF Aspose.Cells | diagonal WordArt watermark Aspose.Cells | shape transparency rotation Aspose.Cells | watermark Z‑order behind cells Aspose.Cells | C# Excel to PDF conversion Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel file, adds a 45‑degree rotated, 50% transparent WordArt watermark labeled CONFIDENTIAL to each worksheet, positions the shape behind the cells, and saves the workbook as a PDF using Aspose.Cells for .NET.
class WorkbookToPdfWithOddPageWatermark
{
    static void Main()
    {
        // Paths for input Excel and output PDF
        string excelPath = "input.xlsx";
        string outputPdfPath = "output.pdf";

        try
        {
            // Verify that the source Excel file exists
            if (!File.Exists(excelPath))
                throw new FileNotFoundException($"The Excel file '{excelPath}' was not found.");

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Watermark settings
            string watermarkText = "CONFIDENTIAL";
            double opacity = 0.5;          // 50% transparent
            double rotationAngle = 45;     // Diagonal

            // Apply watermark to each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Add WordArt text effect as watermark
                    Shape shape = sheet.Shapes.AddTextEffect(
                        MsoPresetTextEffect.TextEffect1,
                        watermarkText,
                        "Arial",
                        72,                     // Font size
                        false,                 // Bold
                        false,                 // Italic
                        0,                     // Upper left X
                        0,                     // Upper left Y
                        sheet.Cells.MaxDisplayRange.ColumnCount * 100, // Width
                        sheet.Cells.MaxDisplayRange.RowCount * 20,    // Height
                        0,                     // Depth (no extrusion)
                        0);                    // Rotation (will set later)

                    // Set rotation and transparency
                    shape.RotationAngle = rotationAngle;
                    shape.Fill.Transparency = opacity;
                    shape.Line.Transparency = opacity;

                    // Optionally, place the shape behind cells by setting Z-order to the lowest value
                    shape.ZOrderPosition = 0;
                }
                catch (Exception exShape)
                {
                    Console.WriteLine($"Warning: Failed to add watermark to sheet '{sheet.Name}'. {exShape.Message}");
                }
            }

            // Save the workbook as PDF
            PdfSaveOptions pdfOptions = new PdfSaveOptions();
            workbook.Save(outputPdfPath, pdfOptions);

            Console.WriteLine($"PDF generated successfully at '{outputPdfPath}'.");
        }
        catch (Exception ex)
        {
            // Log any errors that occur during processing
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
