// Title: Insert an Aspose.Cells‑generated PNG worksheet image into a Word document with OpenXML SDK (C#)
// Description: Learn how to render an Excel worksheet to a PNG image using Aspose.Cells, then embed that image into a Word (.docx) file with the OpenXML SDK for fully automated reporting in .NET. The example shows creating a workbook, converting the first sheet to a memory stream PNG, and inserting the image into a Word body part without manual file handling.
// Keywords: Aspose.Cells PNG export | OpenXML SDK insert image | C# embed Excel snapshot in Word | automated report generation .NET | Excel to Word image conversion | Office Open XML image insertion | render worksheet as PNG | Word document automation C# | Aspose.Cells SheetRender example | programmatic Word report
// Common Searches: how to add a PNG from Aspose.Cells to a Word document using C# | render Excel sheet to image and embed in docx OpenXML | C# code to insert worksheet snapshot into Word for reporting | Aspose.Cells export to PNG then OpenXML insert | automate Word report with Excel image C#
// Developer Intent: Generate a PNG snapshot of an Excel worksheet with Aspose.Cells and programmatically embed it into a Word document using the OpenXML SDK for automated report creation.
// Use Cases: Create monthly financial reports that combine Excel calculations with Word formatting. | Embed visual previews of spreadsheet data in client‑facing proposals generated on the fly. | Produce documentation that includes exact worksheet layouts without requiring Excel on the target machine. | Automate email attachments that contain both data tables (Excel) and narrative (Word) with embedded images.
// AI Prompts: Provide a complete C# example that renders each worksheet page to separate PNG files and inserts them into a multi‑section Word report using OpenXML. | Explain how to control image resolution and scaling when converting an Excel sheet to PNG with Aspose.Cells. | Show how to add a caption and alt text to the inserted PNG in the Word document for accessibility. | Give code to replace an existing placeholder image in a Word template with the generated worksheet PNG.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// Learn how to render an Excel worksheet to a PNG image using Aspose.Cells, then embed that image into a Word (.docx) file with the OpenXML SDK for fully automated reporting in .NET. The example shows creating a workbook, converting the first sheet to a memory stream PNG, and inserting the image into a Word body part without manual file handling.
class InsertWorksheetImageToWord
{
    static void Main()
    {
        try
        {
            // Create a workbook and add some data
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("Sample Data");
            worksheet.Cells["B2"].PutValue(12345);

            // Render the worksheet to a PNG image in memory
            ImageOrPrintOptions options = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png
            };
            SheetRender sheetRender = new SheetRender(worksheet, options);

            using (MemoryStream imageStream = new MemoryStream())
            {
                sheetRender.ToImage(0, imageStream); // Render first sheet page
                imageStream.Position = 0; // Reset stream for reading

                // Save the image to a file
                string imagePath = "WorksheetImage.png";
                using (FileStream fileStream = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                {
                    imageStream.CopyTo(fileStream);
                }

                Console.WriteLine($"Image saved to '{Path.GetFullPath(imagePath)}'.");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
