// Title: C# – Aspose.Cells: Add a Linked Image to Cell A1 and Export as XLSX
// Description: Loads an existing workbook, accesses the first worksheet, inserts a linked picture (100 × 100 px) anchored to cell A1, and saves the result as an XLSX document using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | AddLinkedPicture | linked image Excel C# | anchor picture cell A1 | save workbook XLSX | Shapes API | dynamic logo Excel | C# Excel image linking
// Common Searches: Aspose.Cells add linked picture C# | Insert image into Excel cell programmatically .NET | How to anchor a picture to a specific cell with Aspose.Cells | Save workbook after adding shape Aspose.Cells | C# code to link external image in Excel
// Developer Intent: Insert a linked picture at cell A1 of an existing workbook and generate a new XLSX file.
// Use Cases: Automatic logo updates when the source file changes | Template‑driven reports with placeholder images | Invoices that display product photos linked to external files | Dashboards that pull the latest charts from image files
// AI Prompts: Generate C# code to place a linked picture in cell B2 with width 150 px and height 80 px using Aspose.Cells. | Show how to replace an existing linked picture while keeping its anchor cell. | Explain error handling for missing image files when calling AddLinkedPicture. | Provide a method to batch‑insert linked images into a range of cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads an existing workbook, accesses the first worksheet, inserts a linked picture (100 × 100 px) anchored to cell A1, and saves the result as an XLSX document using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Path to the image that will be linked (replace with your actual image path)
        string imagePath = "sample.jpg";

        // Add a linked picture anchored to cell A1 (row index 0, column index 0)
        // Height and width are specified in pixels
        worksheet.Shapes.AddLinkedPicture(0, 0, 100, 100, imagePath);

        // Save the modified workbook as XLSX
        workbook.Save("output.xlsx", SaveFormat.Xlsx);
    }
}
