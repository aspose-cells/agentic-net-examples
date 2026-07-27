// Title: Add a Linked Picture to Cell A1 and Save as XLSX with Aspose.Cells for .NET (C#)
// Description: Loads input.xlsx, inserts a linked image (sample.jpg) anchored to cell A1 with 100 × 100 px dimensions, and saves the workbook as output.xlsx in XLSX format using Aspose.Cells for .NET.
// Keywords: Aspose.Cells | AddLinkedPicture | linked picture | C# Excel image | anchor image to cell | save workbook as XLSX | Excel shape insertion | Aspose.Cells .NET
// Common Searches: Aspose.Cells add linked picture to cell | C# insert image into Excel cell A1 | anchor picture to specific cell with Aspose.Cells | save workbook after adding picture Aspose.Cells | AddLinkedPicture example C#
// Developer Intent: Insert a linked image into a specific cell of an existing Excel file and produce a new XLSX file using Aspose.Cells for .NET.
// Use Cases: Generate a template where a company logo is linked to a cell, allowing automatic updates when the logo file changes. | Create invoices that pull product photos from external files, keeping the workbook lightweight. | Build a live dashboard that displays external chart images linked to cells, enabling refresh without reopening the workbook.
// AI Prompts: Write C# code to add a linked picture to cell B2 with custom width 150 and height 80 using Aspose.Cells. | Show error handling for a missing image file when calling AddLinkedPicture in Aspose.Cells. | Explain how to position a linked picture over a merged cell range with Aspose.Cells.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Loads input.xlsx, inserts a linked image (sample.jpg) anchored to cell A1 with 100 × 100 px dimensions, and saves the workbook as output.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load an existing workbook
        string inputFile = "input.xlsx";
        Workbook workbook = new Workbook(inputFile);

        // Get the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Path to the image that will be linked
        string imagePath = "sample.jpg";

        // Add a linked picture anchored to cell A1 (row 0, column 0)
        // Height and width are set to 100 pixels each
        worksheet.Shapes.AddLinkedPicture(0, 0, 100, 100, imagePath);

        // Save the workbook as XLSX
        string outputFile = "output.xlsx";
        workbook.Save(outputFile, SaveFormat.Xlsx);
    }
}
