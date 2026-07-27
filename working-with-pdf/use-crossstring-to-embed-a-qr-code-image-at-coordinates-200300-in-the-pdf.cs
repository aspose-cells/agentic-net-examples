using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Author: Aspose.Cells .NET example – embed a QR code image in a PDF

        // 1. Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // (Optional) Add some placeholder text
        sheet.Cells["A1"].PutValue("QR Code Demo");

        // 2. Add the QR code image to the worksheet.
        //    The picture is added at the cell position (row 10, column 5).
        //    Adjust these indices to approximate the desired PDF coordinates (200,300).
        //    The image file "qr.png" must exist in the executable's working directory.
        PictureCollection pictures = sheet.Pictures;
        pictures.Add(10, 5, "qr.png");   // topRow = 10, leftColumn = 5

        // 3. Save the workbook as a PDF. The picture will be rendered at the
        //    location defined above, which corresponds to the requested coordinates.
        workbook.Save("QrCodeOutput.pdf", SaveFormat.Pdf);
    }
}