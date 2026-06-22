using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Rendering;

class EmbedAttachmentsPdfDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            worksheet.Cells["A1"].PutValue("PDF with multiple embedded attachments");

            // -------------------------------------------------
            // Add an image to the worksheet (will appear in PDF)
            // -------------------------------------------------
            string imageFile = "sampleImage.png";

            // Create a simple placeholder PNG if it does not exist (1x1 transparent pixel)
            if (!File.Exists(imageFile))
            {
                const string base64Png = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+XK2cAAAAASUVORK5CYII=";
                byte[] pngBytes = Convert.FromBase64String(base64Png);
                File.WriteAllBytes(imageFile, pngBytes);
            }

            // Insert the image into the worksheet
            worksheet.Pictures.Add(5, 0, imageFile);

            // -------------------------------------------------
            // Create a CSV file and embed it as an OLE object
            // -------------------------------------------------
            string csvFile = "sampleData.csv";
            File.WriteAllText(csvFile, "Name,Score\nAlice,85\nBob,92");

            // Ensure the CSV file exists before reading its bytes
            if (!File.Exists(csvFile))
                throw new FileNotFoundException($"CSV file not found: {csvFile}");

            // Add the CSV file as an OLE object and display it as an icon
            int oleIndex = worksheet.OleObjects.Add(10, 0, 200, 100, File.ReadAllBytes(csvFile));
            worksheet.OleObjects[oleIndex].FileFormatType = FileFormatType.Csv;
            worksheet.OleObjects[oleIndex].DisplayAsIcon = true;

            // -------------------------------------------------
            // Configure PDF save options to embed attachments
            // -------------------------------------------------
            PdfSaveOptions pdfOptions = new PdfSaveOptions
            {
                EmbedAttachments = true // Enable embedding of OLE attachments
            };

            // Save the workbook as a PDF with embedded attachments
            workbook.Save("OutputWithAttachments.pdf", pdfOptions);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
        finally
        {
            // Clean up temporary files if they exist
            string[] tempFiles = { "sampleImage.png", "sampleData.csv" };
            foreach (var file in tempFiles)
            {
                try
                {
                    if (File.Exists(file))
                        File.Delete(file);
                }
                catch
                {
                    // Ignore cleanup errors
                }
            }
        }
    }
}