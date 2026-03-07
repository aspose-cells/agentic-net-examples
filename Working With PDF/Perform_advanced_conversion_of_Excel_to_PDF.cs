using System;
using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Cells.Saving;
using Aspose.Cells.Rendering;

namespace AsposeCellsAdvancedPdfConversion
{
    class Program
    {
        static void Main()
        {
            string sourcePath = "input.xlsx";
            string outputPath = "output.pdf";

            ConversionUtility.Convert(sourcePath, outputPath);
            Console.WriteLine("Simple conversion completed.");

            Workbook workbook = new Workbook(sourcePath);

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                EmbedStandardWindowsFonts = true,
                Compliance = PdfCompliance.PdfA1b,
                CalculateFormula = true,
                OnePagePerSheet = true,
                GridlineType = GridlineType.Dotted,
                Watermark = new RenderingWatermark("CONFIDENTIAL", new RenderingFont("Arial", 72))
                {
                    HAlignment = TextAlignmentType.Center,
                    VAlignment = TextAlignmentType.Center,
                    Rotation = 45,
                    Opacity = 0.3f,
                    ScaleToPagePercent = 50
                }
            };

            workbook.Save(outputPath, pdfOptions);
            Console.WriteLine("Advanced conversion with PDF options completed.");
        }
    }
}