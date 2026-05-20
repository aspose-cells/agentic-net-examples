using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Slicers;

class ConvertWorkbookWithSlicersToPdf
{
    static void Main()
    {
        const string inputPath = "InputWithSlicers.xlsx";
        const string outputPath = "OutputWithSlicers.pdf";

        try
        {
            // Verify the input file exists before loading
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            // Load the workbook containing slicers
            Workbook workbook = new Workbook(inputPath);

            // Ensure all slicers are printable so they appear in the PDF
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                foreach (Slicer slicer in sheet.Slicers)
                {
                    slicer.IsPrintable = true;
                }
            }

            // Configure PDF save options (retain document structure)
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                ExportDocumentStructure = true
            };

            // Save the workbook as PDF
            workbook.Save(outputPath, pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}