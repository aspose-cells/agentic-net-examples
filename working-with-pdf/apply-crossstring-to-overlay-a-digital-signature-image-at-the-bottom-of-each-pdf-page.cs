// Title: Add a digital signature image to the footer of each PDF page when converting Excel to PDF with Aspose.Cells for .NET
// Description: C# sample that loads an Excel workbook, verifies a PNG signature file, inserts the image a few rows below the last used row on every worksheet, scales it to the width of ten columns while keeping the original aspect ratio, and saves the workbook as a PDF so the signature appears in the footer of each generated page.
// Keywords: Aspose.Cells add image | digital signature footer PDF | Excel to PDF with picture | C# Aspose.Cells image scaling | place image at bottom of worksheet | preserve aspect ratio Aspose.Cells | PDF footer image automation
// Common Searches: how to add a signature image to each PDF page using Aspose.Cells | Aspose.Cells insert picture at worksheet bottom before PDF export | scale inserted image to column width in Aspose.Cells C# | preserve image aspect ratio when adding footer in Excel to PDF conversion | automate PDF footer logo with Aspose.Cells
// Developer Intent: Insert a digital signature (or any footer image) into every worksheet so it appears at the bottom of each page when the workbook is saved as a PDF.
// Use Cases: Brand every exported report with a company logo in the PDF footer. | Attach a handwritten signature to invoices before PDF generation for compliance. | Add a security seal or certification badge to all pages of a generated statement.
// AI Prompts: Write C# code using Aspose.Cells to place an image a few rows below the last data row on each worksheet, fit it to ten column widths, keep its aspect ratio, and export the workbook to PDF. | Explain how to calculate printable width from column widths in Aspose.Cells and apply it to a footer picture. | Provide error‑handling patterns for missing Excel files or signature images when adding a footer image before PDF conversion.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// C# sample that loads an Excel workbook, verifies a PNG signature file, inserts the image a few rows below the last used row on every worksheet, scales it to the width of ten columns while keeping the original aspect ratio, and saves the workbook as a PDF so the signature appears in the footer of each generated page.
class Program
{
    static void Main()
    {
        try
        {
            // Path to the source workbook
            string sourcePath = "input.xlsx";

            // Verify that the source file exists
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(sourcePath);

            // Path to the digital signature image
            string signatureImagePath = "signature.png";

            // Verify that the signature image exists
            if (!File.Exists(signatureImagePath))
            {
                Console.WriteLine($"Signature image not found: {signatureImagePath}");
                return;
            }

            // Add the signature image to the bottom of each worksheet
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                // Determine a row near the bottom of the used range
                int lastUsedRow = sheet.Cells.MaxDataRow;
                int targetRow = lastUsedRow + 2; // a couple of rows below the data

                // Insert the picture anchored to the calculated cell (column 0 = A)
                int pictureIndex = sheet.Pictures.Add(targetRow, 0, signatureImagePath);
                Picture picture = sheet.Pictures[pictureIndex];

                // Approximate printable width using the width of several columns (e.g., 10 columns)
                int printableWidth = sheet.Cells.GetColumnWidthPixel(0) * 10;
                picture.Width = printableWidth;

                // Preserve aspect ratio
                double aspectRatio = (double)picture.OriginalHeight / picture.OriginalWidth;
                picture.Height = (int)(picture.Width * aspectRatio);
            }

            // Save the workbook as PDF; the signature image will appear at the bottom of each page
            string outputPath = "output.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF: {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
