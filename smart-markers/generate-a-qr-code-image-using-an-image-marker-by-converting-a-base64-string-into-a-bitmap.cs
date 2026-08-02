// Title: Insert QR Code from Base64 into Excel with Aspose.Cells (C#)
// Description: C# sample that decodes a Base64‑encoded PNG, streams it into a worksheet cell via Pictures.Add, and saves the workbook as an .xlsx file.
// Keywords: Aspose.Cells | C# QR code | Base64 image Excel | Add picture to worksheet | MemoryStream image | Insert QR code Excel | Convert Base64 to bitmap | Image marker Aspose.Cells
// Common Searches: Aspose.Cells add image from Base64 | Insert QR code into Excel C# | Convert Base64 PNG to bitmap Aspose.Cells | Place picture in specific cell using memory stream | Smart marker image example Aspose.Cells
// Developer Intent: Add a Base64‑encoded QR‑code image to a specific cell in an Excel file using Aspose.Cells.
// Use Cases: Generate sales reports that embed a QR code for each transaction. | Create product‑label worksheets where QR codes are supplied as Base64 strings. | Automate invoice generation with a QR‑code signature placed in a designated cell.
// AI Prompts: Write C# code with Aspose.Cells to decode a Base64 PNG and insert it at row 5, column 3. | Show how to loop through a list of Base64 QR codes and add each to consecutive rows in an Excel sheet. | Explain error handling for invalid Base64 data when adding an image to an Aspose.Cells worksheet.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsQrCodeDemo
{
    // C# sample that decodes a Base64‑encoded PNG, streams it into a worksheet cell via Pictures.Add, and saves the workbook as an .xlsx file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Base64 string of a 1x1 pixel PNG image (valid Base64 data)
                const string qrCodeBase64 =
                    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+X9WcAAAAASUVORK5CYII=";

                // Convert Base64 string to byte array
                byte[] imageBytes = Convert.FromBase64String(qrCodeBase64);

                // Add the QR code image to the worksheet using a memory stream
                using (MemoryStream ms = new MemoryStream(imageBytes))
                {
                    // Row and column indices where the image will be placed (0‑based)
                    int row = 2;
                    int column = 2;
                    sheet.Pictures.Add(row, column, ms);
                }

                // Save the workbook to a file
                const string outputPath = "QrCodeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}
