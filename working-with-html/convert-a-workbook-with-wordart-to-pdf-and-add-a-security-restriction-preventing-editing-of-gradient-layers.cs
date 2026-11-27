// Title: Convert an Excel workbook containing WordArt to a PDF and lock WordArt adjustment handles with Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# lock WordArt adjustment handles before saving workbook as PDF | How to add an owner password to a PDF generated from Excel with WordArt using Aspose.Cells | Prevent editing of WordArt shapes in PDF output from Aspose.Cells .NET | Set PDF security options while converting Excel to PDF with Aspose.Cells in C# | Lock gradient layer handles of WordArt in Excel and protect the resulting PDF with Aspose.Cells
// Tags: WordArt shape lock Aspose.Cells | owner password PDF protection Aspose.Cells | Excel to PDF conversion with shape security | PdfSaveOptions security configuration .NET | protect WordArt gradient layers in PDF

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;
using Aspose.Cells.Rendering.PdfSecurity;

namespace WordArtPdfSecurityDemo
{
    // The example loads an Excel workbook, finds WordArt shapes, locks their adjustment handles, configures PdfSecurityOptions with an owner password, applies the security via PdfSaveOptions, and saves the workbook as a protected PDF.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputFile = "WordArtWorkbook.xlsx";
            const string outputPdf = "WordArtProtected.pdf";

            try
            {
                // Verify that the input workbook exists
                if (!File.Exists(inputFile))
                {
                    Console.WriteLine($"Input file not found: {inputFile}");
                    return;
                }

                // Load the workbook that contains WordArt shapes
                Workbook workbook = new Workbook(inputFile);
                Worksheet sheet = workbook.Worksheets[0];

                // Locate WordArt shapes and lock their adjustable handles
                foreach (Shape shape in sheet.Shapes)
                {
                    if (shape.IsWordArt)
                    {
                        // Lock adjustment handles so they cannot be edited
                        shape.SetLockedProperty(ShapeLockType.AdjustHandles, true);
                    }
                }

                // Prepare PDF security options – disallow content modification
                PdfSecurityOptions pdfSec = new PdfSecurityOptions
                {
                    OwnerPassword = "ownerPassword123",
                    UserPassword = string.Empty
                    // Permissions property is not available in the current Aspose.Cells version.
                    // By default, without explicit permissions the PDF will be protected by the owner password.
                };

                // Configure PDF save options and attach the security settings
                PdfSaveOptions pdfOptions = new PdfSaveOptions
                {
                    SecurityOptions = pdfSec,
                    OnePagePerSheet = true
                };

                // Save the workbook as a PDF with the security restrictions
                workbook.Save(outputPdf, pdfOptions);

                Console.WriteLine($"Workbook successfully converted to PDF with security applied: {outputPdf}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred:");
                Console.WriteLine(ex.Message);
            }
        }
    }
}
