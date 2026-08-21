// Title: Export Modified SmartArt Shapes to PDF with Aspose.Cells for .NET
// Description: Loads an Excel workbook, iterates through all worksheets and shapes, updates each shape's alternative text, converts SmartArt shapes to GroupShape objects, modifies the text of inner shapes, configures PdfSaveOptions to refresh SmartArt and preserve document structure, and saves the result as a PDF.
// Keywords: Aspose.Cells | C# | SmartArt export to PDF | PdfSaveOptions UpdateSmartArt | GroupShape manipulation | Excel to PDF conversion | modify shape alternative text | preserve PDF document structure
// Common Searches: Aspose.Cells update SmartArt before PDF export | C# convert SmartArt to GroupShape and save as PDF | PdfSaveOptions.UpdateSmartArt example | change shape alternative text in Excel with Aspose.Cells | export Excel workbook with modified SmartArt to PDF
// Developer Intent: Generate a PDF from an Excel file after programmatically editing SmartArt shapes using Aspose.Cells for .NET.
// Use Cases: Refresh all SmartArt graphics after changing their text or properties and produce an up‑to‑date PDF. | Batch‑process workbooks to replace alternative text on every shape and ensure the changes appear in the exported PDF. | Maintain the logical document structure in the PDF while applying SmartArt modifications for accessibility compliance.
// AI Prompts: Provide C# code that loads an .xlsx file, updates SmartArt node text, and saves the workbook as a PDF with UpdateSmartArt enabled. | Explain how PdfSaveOptions.UpdateSmartArt and ExportDocumentStructure affect the PDF output when using Aspose.Cells. | Show how to iterate through shapes, detect SmartArt, convert it to a GroupShape, and modify inner shapes in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an Excel workbook, iterates through all worksheets and shapes, updates each shape's alternative text, converts SmartArt shapes to GroupShape objects, modifies the text of inner shapes, configures PdfSaveOptions to refresh SmartArt and preserve document structure, and saves the result as a PDF.
public class ExportSmartArtToPdf
{
    public static void Main(string[] args)
    {
        try
        {
            Run();
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    public static void Run()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.pdf";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        // Load the source workbook
        Workbook workbook = new Workbook(inputPath);

        // Iterate through all worksheets and their shapes
        foreach (Worksheet worksheet in workbook.Worksheets)
        {
            foreach (Shape shape in worksheet.Shapes)
            {
                // Example modification: change the alternative text of every shape
                shape.AlternativeText = "ModifiedAltText";

                // Process only SmartArt shapes
                if (shape.IsSmartArt)
                {
                    // Convert the SmartArt to a grouped shape collection
                    GroupShape groupShape = shape.GetResultOfSmartArt();

                    if (groupShape != null)
                    {
                        // Modify each inner shape within the SmartArt group
                        foreach (Shape innerShape in groupShape.GetGroupedShapes())
                        {
                            // Example modification: set new text for each inner shape
                            innerShape.Text = "ModifiedSmartArtText";
                        }
                    }
                }
            }
        }

        // Configure PDF save options and enable SmartArt updating
        PdfSaveOptions pdfOptions = new PdfSaveOptions
        {
            UpdateSmartArt = true,               // Ensure SmartArt changes are reflected
            ExportDocumentStructure = true       // Optional: retain document structure in PDF
        };

        // Save the modified workbook as a PDF using the specified options
        workbook.Save(outputPath, pdfOptions);

        Console.WriteLine($"PDF saved successfully to {outputPath}");
    }
}
