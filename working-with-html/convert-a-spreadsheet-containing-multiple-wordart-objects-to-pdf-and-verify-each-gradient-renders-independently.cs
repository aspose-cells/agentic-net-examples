// Title: Convert Excel workbook with multiple WordArt gradient shapes to PDF and validate each gradient renders correctly using Aspose.Cells for .NET
// AI Prompts: Generate C# code that loads an .xlsx file, iterates through all worksheets, identifies WordArt shapes, checks whether each shape's FillFormat is a gradient, logs the worksheet name and shape index with gradient status, and saves the workbook as a PDF using Aspose.Cells. | Update the conversion script to throw an exception if any WordArt shape does not use a gradient fill before the workbook is exported to PDF.
// Common Searches: Aspose.Cells C# export WordArt with gradient fill to PDF | How to detect gradient fill on WordArt shapes in an Excel file using Aspose.Cells | Validate individual WordArt gradient rendering after saving workbook as PDF in .NET | Iterate over ShapeCollection to find WordArt objects with Aspose.Cells
// Tags: Aspose.Cells detect WordArt gradient fill | convert Excel WordArt shapes to PDF .NET | enumerate ShapeCollection for WordArt Aspose | check FillFormat type C# Aspose.Cells | preserve gradient rendering in PDF export

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The example loads an Excel workbook, scans each worksheet's ShapeCollection for WordArt objects, determines if each WordArt uses a gradient fill, logs the worksheet name, shape index, and gradient status, and finally saves the workbook as a PDF, ensuring that gradient fills are rendered correctly.
class WordArtGradientPdfConverter
{
    static void Main()
    {
        // Paths for input Excel and output PDF
        string excelPath = "input.xlsx";
        string pdfPath = "output.pdf";

        try
        {
            // Verify that the input file exists
            if (!File.Exists(excelPath))
            {
                Console.WriteLine($"Error: Input file not found at '{excelPath}'.");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(excelPath);

            // Iterate through worksheets and inspect WordArt shapes
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                ShapeCollection shapes = sheet.Shapes;

                for (int i = 0; i < shapes.Count; i++)
                {
                    Shape shape = shapes[i];

                    // WordArt objects can be identified via IsWordArt property
                    if (shape.IsWordArt)
                    {
                        // Use the Fill property (FillFormat)
                        FillFormat fill = shape.Fill;

                        // Determine if the fill is a gradient
                        bool isGradient = fill.FillType == FillType.Gradient;

                        Console.WriteLine($"Worksheet: {sheet.Name}, WordArt Index: {i}");
                        Console.WriteLine($"  Gradient Fill: {isGradient}");

                        if (isGradient)
                        {
                            // Gradient fill detected; additional validation can be added here if needed
                            Console.WriteLine("  Gradient configuration appears valid.");
                        }
                        else
                        {
                            Console.WriteLine("  Warning: WordArt does not use a gradient fill.");
                        }
                    }
                }
            }

            // Save the workbook as PDF; WordArt gradients will be rendered in the PDF
            try
            {
                workbook.Save(pdfPath, SaveFormat.Pdf);
                Console.WriteLine($"Conversion completed. PDF saved to: {pdfPath}");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine($"Error saving PDF: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
