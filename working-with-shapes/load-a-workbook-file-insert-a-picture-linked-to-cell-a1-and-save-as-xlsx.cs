using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

class Program
{
    static void Main()
    {
        // Paths for the source workbook, the image to link, and the output workbook
        string sourceWorkbookPath = "input.xlsx";
        string linkedImagePath = "image.jpg";
        string outputWorkbookPath = "output.xlsx";

        // Load the existing workbook
        Workbook workbook = new Workbook(sourceWorkbookPath);

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Insert a linked picture anchored to cell A1 (row 0, column 0)
        // Height and width are specified in pixels (e.g., 100x100)
        worksheet.Shapes.AddLinkedPicture(0, 0, 100, 100, linkedImagePath);

        // Save the modified workbook as XLSX
        workbook.Save(outputWorkbookPath, SaveFormat.Xlsx);
    }
}