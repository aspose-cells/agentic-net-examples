// Title: Insert a QR‑code image from a Base64 string into an Excel cell with Aspose.Cells for .NET
// Description: Demonstrates how to decode a Base64‑encoded QR‑code PNG, load it into a MemoryStream, and place the image in a specific worksheet cell using Aspose.Cells (C#) before saving as XLSX.
// Keywords: Aspose.Cells | C# | Base64 to image | QR code Excel | Worksheet.Pictures.Add | MemoryStream image | Insert picture into cell | .NET Excel image marker
// Common Searches: Aspose.Cells add image from Base64 | C# embed QR code in Excel worksheet | How to use Worksheet.Pictures.Add with a stream | Convert Base64 string to bitmap for Excel | Insert PNG picture into specific cell Aspose.Cells
// Developer Intent: Place a QR‑code image decoded from a Base64 string into a designated cell of an Excel workbook using Aspose.Cells.
// Use Cases: Generate sales reports that embed a QR code for each record. | Create printable product label sheets with QR codes directly in Excel. | Automate invoice PDFs by inserting QR‑code payment links as images in the workbook.
// AI Prompts: Write C# code with Aspose.Cells to decode a Base64 QR‑code and insert it at cell C5. | Show how to resize the inserted QR‑code picture and apply a border using Aspose.Cells. | Provide an example that loops through a list of Base64 QR‑code strings and adds each image to successive rows in an Excel sheet.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsQrCodeDemo
{
    // Demonstrates how to decode a Base64‑encoded QR‑code PNG, load it into a MemoryStream, and place the image in a specific worksheet cell using Aspose.Cells (C#) before saving as XLSX.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Base64 string representing a QR code PNG image.
            // Replace this string with the actual Base64 data of your QR code.
            const string qrCodeBase64 = "iVBORw0KGgoAAAANSUhEUgAAAOEAAADhCAYAAAB6V+0UAAAABHNCSVQICAgIfAhkiAAAAAlwSFlzAAAOxAAADsQBlSsOGwAAABl0RVh0U29mdHdhcmUAcGFpbnQubmV0IDQuMi4xMZ8Z3wAAABl0RVh0Q3JlYXRpb24gVGltZQAwOS8yMi8xM5Z6+QAAABV0RVh0U291cmNlAEFzcG9zZSBJbWFnZSBMaWJyYXJ5IEV4YW1wbGUgQ1JYAAAAAElFTkSuQmCC";

            // Convert the Base64 string to a byte array
            byte[] qrCodeBytes = Convert.FromBase64String(qrCodeBase64);

            // Load the byte array into a memory stream
            using (MemoryStream imageStream = new MemoryStream(qrCodeBytes))
            {
                // Add the QR code image to the worksheet at cell B2 (row 1, column 1)
                // The Pictures.Add method accepts a stream containing image data.
                sheet.Pictures.Add(1, 1, imageStream);
            }

            // Save the workbook to an XLSX file
            workbook.Save("QrCodeWorkbook.xlsx");
        }
    }
}
