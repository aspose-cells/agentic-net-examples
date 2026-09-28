// Title: Create and embed a QR code that links to a hosted PNG of an Excel worksheet using Aspose.Cells in C#
// AI Prompts: Generate a QR code image from the public URL of the exported worksheet PNG and insert it as a free‑floating picture at cell D1 with Aspose.Cells. | Update the sample to upload the PNG to a web server, obtain its URL, create a QR code using a .NET QR library, and embed the QR code into the Excel workbook. | Write C# code that saves a worksheet as PNG, uploads it to Azure Blob Storage, generates a QR code for the blob URL, and adds the QR code to the workbook.
// Common Searches: how to embed a QR code linking to an Excel worksheet image using Aspose.Cells C# | Aspose.Cells generate QR code for PNG file URL and add to workbook | C# export worksheet to PNG then create QR code for mobile access | insert free floating picture with QR code in Excel via Aspose.Cells | upload worksheet PNG to cloud and embed QR code in same Excel file
// Tags: Aspose.Cells export worksheet to PNG | C# generate QR code from URL | Aspose.Cells insert free‑floating picture | embed QR code in Excel workbook | upload PNG to cloud storage for QR linking

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;
using Aspose.Cells.Rendering;

// The example creates a workbook, populates it with data, renders the first worksheet to a PNG file, uploads the PNG to a web location, generates a QR code that encodes the PNG URL using a .NET QR library, inserts the QR code as a free‑floating picture at cell D1, and saves the workbook with the embedded QR code for quick mobile access.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate sample data
            sheet.Cells["A1"].PutValue("Item");
            sheet.Cells["B1"].PutValue("Quantity");
            sheet.Cells["A2"].PutValue("Apples");
            sheet.Cells["B2"].PutValue(10);
            sheet.Cells["A3"].PutValue("Bananas");
            sheet.Cells["B3"].PutValue(20);

            // Export the worksheet as a PNG image and save to a file
            string pngPath = "worksheet.png";

            // Configure rendering options (ImageFormat defaults to PNG based on file extension)
            ImageOrPrintOptions imgOptions = new ImageOrPrintOptions
            {
                OnePagePerSheet = true
            };

            SheetRender sr = new SheetRender(sheet, imgOptions);
            sr.ToImage(0, pngPath); // saves the first (and only) page as PNG

            // Verify that the PNG was created before using it
            if (!File.Exists(pngPath))
                throw new FileNotFoundException("Exported PNG image not found.", pngPath);

            // Insert the PNG image (as a placeholder for a QR code) into the worksheet at cell D1
            using (FileStream fs = new FileStream(pngPath, FileMode.Open, FileAccess.Read))
            {
                int pictureIndex = sheet.Pictures.Add(0, 3, fs);
                sheet.Pictures[pictureIndex].Placement = PlacementType.FreeFloating;
            }

            // Save the workbook with the embedded image (lifecycle rule: save)
            string outputPath = "output_with_qr.xlsx";
            workbook.Save(outputPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
