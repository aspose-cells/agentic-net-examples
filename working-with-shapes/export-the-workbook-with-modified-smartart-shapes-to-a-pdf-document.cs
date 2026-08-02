// Title: Export Modified SmartArt Shapes to PDF with Aspose.Cells for .NET (C#)
// Description: Loads an .xlsx file, iterates through all worksheets and shapes, updates each shape's alternative text, converts SmartArt to grouped shapes to edit inner text, sets PdfSaveOptions (UpdateSmartArt and ExportDocumentStructure), and saves the workbook as a PDF with the changes reflected.
// Keywords: Aspose.Cells C# PDF export | SmartArt to PDF Aspose | UpdateSmartArt option | modify shape alternative text | convert SmartArt to group shapes | PdfSaveOptions ExportDocumentStructure | Excel to PDF Aspose .NET | Aspose.Cells shape manipulation
// Common Searches: Aspose.Cells export SmartArt to PDF C# | How to update SmartArt text before PDF conversion Aspose | PdfSaveOptions UpdateSmartArt example | Convert SmartArt to grouped shapes Aspose.Cells | Save Excel as PDF with document structure Aspose
// Developer Intent: Create a PDF from an Excel workbook after programmatically modifying SmartArt and other shape attributes using Aspose.Cells for .NET.
// Use Cases: Refresh all shape alternative texts and produce a PDF that reflects the new metadata. | Break down each SmartArt object into its component shapes, change their displayed text, and export the workbook while preserving the updated SmartArt layout. | Generate a PDF that retains the workbook's document structure and ensures SmartArt graphics are re‑rendered with the latest changes.
// AI Prompts: Write C# code that loads an .xlsx file, changes the AlternativeText of every shape, updates the text of SmartArt components, and saves the workbook as a PDF using PdfSaveOptions.UpdateSmartArt. | Explain how the PdfSaveOptions.UpdateSmartArt flag influences SmartArt rendering during PDF conversion in Aspose.Cells. | Show how to retrieve grouped shapes from a SmartArt object with GetResultOfSmartArt() and modify each shape's Text property.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

namespace AsposeCellsSmartArtPdfExport
{
    // Loads an .xlsx file, iterates through all worksheets and shapes, updates each shape's alternative text, converts SmartArt to grouped shapes to edit inner text, sets PdfSaveOptions (UpdateSmartArt and ExportDocumentStructure), and saves the workbook as a PDF with the changes reflected.
    public class ExportSmartArtToPdf
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.pdf";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
                return;
            }

            try
            {
                // Load the workbook that contains SmartArt shapes
                Workbook workbook = new Workbook(inputPath);

                // Iterate through all worksheets and their shapes
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    foreach (Shape shape in sheet.Shapes)
                    {
                        // Example modification: change alternative text for every shape
                        shape.AlternativeText = "ModifiedAltText";

                        // If the shape is a SmartArt, convert it to grouped shapes and modify the text
                        if (shape.IsSmartArt)
                        {
                            // Convert SmartArt to a group of shapes
                            GroupShape group = shape.GetResultOfSmartArt();

                            // Iterate through each shape inside the group and change its text
                            foreach (Shape smartArtShape in group.GetGroupedShapes())
                            {
                                smartArtShape.Text = "ModifiedSmartArtText";
                            }
                        }
                    }
                }

                // Configure PDF save options to update SmartArt during saving
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    UpdateSmartArt = true,               // Ensure SmartArt is refreshed
                    ExportDocumentStructure = true       // Retain document structure in PDF (optional)
                };

                // Save the modified workbook as a PDF
                workbook.Save(outputPath, pdfOptions);
                Console.WriteLine($"PDF successfully saved to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors during processing
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            ExportSmartArtToPdf.Run();
        }
    }
}
