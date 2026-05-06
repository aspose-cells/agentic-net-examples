using System;
using System.IO;
using Aspose.Cells;

namespace ExternalResourcePdfDemo
{
    public class FileStreamProvider : IStreamProvider
    {
        public void InitStream(StreamProviderOptions options)
        {
            options.Stream = new FileStream(options.DefaultPath, FileMode.Open, FileAccess.Read);
        }

        public void CloseStream(StreamProviderOptions options)
        {
            if (options.Stream != null)
            {
                options.Stream.Close();
                options.Stream = null;
            }
        }
    }

    class Program
    {
        static void Main()
        {
            // Ensure a sample image exists (1x1 pixel PNG).
            string imagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "sample.png");
            if (!File.Exists(imagePath))
            {
                // Transparent 1x1 PNG.
                byte[] pngData = Convert.FromBase64String(
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X2ZcAAAAASUVORK5CYII=");
                File.WriteAllBytes(imagePath, pngData);
            }

            Workbook workbook = new Workbook();

            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("External Resource PDF Demo");
            sheet.Cells["A2"].PutValue("Image loaded via ResourceProvider");

            // Insert picture using the existing image file.
            sheet.Pictures.Add(2, 0, imagePath);

            workbook.Settings.DisplayDrawingObjects = DisplayDrawingObjects.DisplayShapes;
            workbook.Settings.ResourceProvider = new FileStreamProvider();

            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true,
                EmbedAttachments = true,
                DefaultFont = "Arial",
                CheckFontCompatibility = true,
                OnePagePerSheet = true
            };

            string outputPdf = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExternalResourceDemo.pdf");
            workbook.Save(outputPdf, pdfOptions);

            Console.WriteLine($"PDF saved to: {outputPdf}");
        }
    }
}