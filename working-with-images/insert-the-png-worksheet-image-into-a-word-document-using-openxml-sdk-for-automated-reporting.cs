// Title: Insert Excel Worksheet PNG into a Word Document with OpenXML SDK using Aspose.Cells (C#)
// Description: A C# walkthrough that renders an Excel worksheet to a PNG image with Aspose.Cells, then embeds the image into a .docx file via the OpenXML SDK, providing a ready‑to‑use solution for automated reporting.
// Keywords: Aspose.Cells render worksheet PNG | OpenXML SDK insert picture Word | C# embed Excel image into Word document | automated report generation Excel to Word | SheetRender PNG output | ImageOrPrintOptions PNG Aspose.Cells | WordprocessingDocument add image | C# create Word report with spreadsheet snapshot
// Common Searches: how to add a PNG from Excel to a Word file using C# | Aspose.Cells render worksheet as image and insert into docx | OpenXML SDK embed picture programmatically | C# generate Word report from Excel data | export Excel sheet to PNG and place in Word document
// Developer Intent: Generate a PNG snapshot of an Excel worksheet and programmatically embed it into a Word document for automated reporting.
// Use Cases: Create visual summaries of spreadsheet data inside corporate Word reports. | Automate the production of client‑facing documents that combine Excel calculations with formatted Word layouts. | Generate PDF or DOCX deliverables where Excel charts or tables must appear as static images.
// AI Prompts: Write C# code that uses Aspose.Cells to render the first worksheet page to a PNG stream and then inserts the image into a new Word document using the OpenXML SDK. | Show how to configure ImageOrPrintOptions for PNG output, create a MemoryStream for the image, and add the picture to a WordprocessingDocument with proper relationship IDs. | Provide an end‑to‑end example that saves the Word file to disk after embedding the worksheet image, handling folder creation and exception management.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

// A C# walkthrough that renders an Excel worksheet to a PNG image with Aspose.Cells, then embeds the image into a .docx file via the OpenXML SDK, providing a ready‑to‑use solution for automated reporting.
class InsertWorksheetImageIntoWord
{
    static void Main()
    {
        try
        {
            // 1. Create a new workbook and add sample data
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample data for image");

            // 2. Render the worksheet to a PNG image in memory
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                ImageType = Aspose.Cells.Drawing.ImageType.Png,
                OnePagePerSheet = true
            };
            SheetRender sheetRender = new SheetRender(sheet, imgOptions);

            // 3. Save the rendered image to a file
            string imagePath = "WorksheetImage.png";
            using (MemoryStream imgStream = new MemoryStream())
            {
                sheetRender.ToImage(0, imgStream);
                imgStream.Position = 0;

                // Ensure the directory exists
                string directory = Path.GetDirectoryName(Path.GetFullPath(imagePath));
                if (!Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                // Write the image to disk
                using (FileStream fileStream = new FileStream(imagePath, FileMode.Create, FileAccess.Write))
                {
                    imgStream.CopyTo(fileStream);
                }
            }

            Console.WriteLine($"Worksheet image saved at: {Path.GetFullPath(imagePath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
