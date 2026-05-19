using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsSmartArtPdfExport
{
    public class ExportSmartArtToPdf
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Ensure the input file exists
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing SmartArt shapes
            Workbook workbook = new Workbook(inputPath);

            // Iterate through all worksheets and shapes
            foreach (Worksheet worksheet in workbook.Worksheets)
            {
                foreach (Shape shape in worksheet.Shapes)
                {
                    // Example modification: change alternative text
                    shape.AlternativeText = "ModifiedAltText";

                    // If the shape is SmartArt, convert it to grouped shapes and modify each part
                    if (shape.IsSmartArt)
                    {
                        GroupShape group = shape.GetResultOfSmartArt();
                        if (group != null)
                        {
                            foreach (Shape smartShape in group.GetGroupedShapes())
                            {
                                smartShape.Text = "ModifiedSmartArtText";
                            }
                        }
                    }
                }
            }

            // Configure PDF save options to update SmartArt and export document structure
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                UpdateSmartArt = true,
                ExportDocumentStructure = true
            };

            // Save the modified workbook as a PDF
            workbook.Save(outputPath, pdfOptions);
        }

        // Entry point for the application
        public static void Main(string[] args)
        {
            try
            {
                Run();
                Console.WriteLine("Export completed successfully.");
            }
            catch (FileNotFoundException ex)
            {
                Console.Error.WriteLine($"File error: {ex.Message}");
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"An unexpected error occurred: {ex.Message}");
            }
        }
    }
}