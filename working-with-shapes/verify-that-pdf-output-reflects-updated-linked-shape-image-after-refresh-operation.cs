// Title: Refresh a linked picture shape and export an updated PDF with Aspose.Cells for .NET (C#)
// Description: Loads an Excel workbook containing a linked picture, calls Picture.UpdateSelectedValue() to refresh the image from its source cell, and saves the workbook as a PDF so the refreshed picture appears in the output. Includes error handling for missing files or shapes.
// Keywords: Aspose.Cells | C# | .NET | linked picture shape | UpdateSelectedValue | refresh linked image | export Excel to PDF | picture shape refresh | Excel automation | PDF generation
// Common Searches: Aspose.Cells refresh linked picture before PDF export | UpdateSelectedValue picture shape C# | How to refresh linked image in Excel using Aspose.Cells | Export workbook to PDF with updated linked picture | C# Aspose.Cells picture shape refresh
// Developer Intent: Refresh the linked picture shape so the generated PDF contains the latest image.
// Use Cases: Automated report pipelines where linked logos or charts must reflect current data before PDF creation. | Batch conversion of multiple workbooks to PDFs, ensuring each linked picture is up‑to‑date. | Generating client‑ready PDFs from spreadsheets that include dynamic images tied to cell values. | Testing workflows that validate visual consistency of linked images after data changes.
// AI Prompts: Show C# code that iterates over all picture shapes in a worksheet, calls UpdateSelectedValue on each, and then saves the workbook as PDF using Aspose.Cells. | Provide robust error handling for missing input files, absent picture shapes, and save failures when refreshing linked images and exporting to PDF. | Explain how to programmatically verify that the PDF contains the refreshed linked picture after calling UpdateSelectedValue. | Suggest a unit‑test approach to confirm that the picture image in the PDF matches the current cell value.

using Aspose.Cells;
using Aspose.Cells.Drawing;
using System;
using System.IO;

// Loads an Excel workbook containing a linked picture, calls Picture.UpdateSelectedValue() to refresh the image from its source cell, and saves the workbook as a PDF so the refreshed picture appears in the output. Includes error handling for missing files or shapes.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "LinkedPicture.xlsx";

            // Verify that the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file '{inputPath}' not found. Operation aborted.");
                return;
            }

            // Load the workbook that contains a linked picture shape
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Locate the first picture shape in the worksheet
            Picture linkedPicture = null;
            foreach (Shape shape in worksheet.Shapes)
            {
                if (shape is Picture picture)
                {
                    linkedPicture = picture;
                    break;
                }
            }

            // Refresh the picture so it reflects the current value of its linked cell
            if (linkedPicture != null)
            {
                linkedPicture.UpdateSelectedValue();
            }
            else
            {
                Console.WriteLine("No picture shape found in the worksheet.");
            }

            // Save the workbook as PDF; the PDF will contain the refreshed picture
            string outputPath = "LinkedPicture_Refreshed.pdf";
            workbook.Save(outputPath, SaveFormat.Pdf);
            Console.WriteLine($"Workbook saved as PDF to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
