using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class SvgWorksheetToPdf
{
    static void Main()
    {
        try
        {
            const string svgFilePath = "image.svg";
            const string pdfOutputPath = "output.pdf";

            // Verify SVG source file exists
            if (!File.Exists(svgFilePath))
            {
                Console.WriteLine($"SVG file not found: {Path.GetFullPath(svgFilePath)}");
                return;
            }

            // Load SVG bytes
            byte[] svgBytes = File.ReadAllBytes(svgFilePath);

            // Create a new workbook and add the SVG shape
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            ShapeCollection shapes = sheet.Shapes;

            // Add SVG at row 4, column 5 (zero‑based indices); size auto‑adjusted
            shapes.AddSvg(4, 0, 5, 0, -1, -1, svgBytes, null);

            // Save the workbook directly to PDF (vector graphics are preserved)
            workbook.Save(pdfOutputPath, SaveFormat.Pdf);

            Console.WriteLine("PDF generated successfully: " + Path.GetFullPath(pdfOutputPath));
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}