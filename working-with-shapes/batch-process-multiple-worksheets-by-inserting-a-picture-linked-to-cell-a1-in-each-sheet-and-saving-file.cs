// Title: C# – Add a linked picture to cell A1 of every worksheet in an Aspose.Cells workbook
// Description: Creates a workbook, ensures multiple worksheets exist, loops through each sheet, and uses Shapes.AddLinkedPicture to insert the same external image at cell A1 (100 × 100 px). The file is saved as BatchLinkedPictures.xlsx.
// Keywords: Aspose.Cells linked picture | AddLinkedPicture C# | insert image all worksheets | batch picture insertion Aspose.Cells | C# Aspose.Cells shape | cell A1 picture | multiple worksheets image
// Common Searches: Aspose.Cells add linked picture to all sheets | C# loop worksheets insert image Aspose.Cells | place same picture in every worksheet using Aspose.Cells | batch add picture cell A1 Aspose.Cells .NET
// Developer Intent: Insert the same external image as a linked picture into cell A1 of each worksheet and save the workbook.
// Use Cases: Add a company logo to the top‑left corner of every sheet for consistent branding. | Apply a confidentiality watermark across all worksheets. | Provide a placeholder graphic in template files so users know where to insert their own images.
// AI Prompts: Write C# code that opens an existing Aspose.Cells workbook, iterates through all worksheets, and adds a linked picture from a specified file path to cell A1 with custom dimensions, then saves the workbook. | Show how to modify the sample so the linked picture is added only when cell A1 is empty, and log the names of sheets where the picture was inserted.

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a workbook, ensures multiple worksheets exist, loops through each sheet, and uses Shapes.AddLinkedPicture to insert the same external image at cell A1 (100 × 100 px). The file is saved as BatchLinkedPictures.xlsx.
class Program
{
    static void Main()
    {
        // Create a new workbook (empty file)
        Workbook workbook = new Workbook();

        // Ensure there are multiple worksheets to process
        workbook.Worksheets[0].Name = "Sheet1";
        workbook.Worksheets.Add("Sheet2");
        workbook.Worksheets.Add("Sheet3");

        // Path to the image that will be linked (adjust as needed)
        string imagePath = "image.jpg";

        // Loop through every worksheet and insert a linked picture at cell A1 (row 0, column 0)
        foreach (Worksheet sheet in workbook.Worksheets)
        {
            // AddLinkedPicture(topRow, leftColumn, height, width, sourceFullName)
            // Height and width are specified in pixels.
            sheet.Shapes.AddLinkedPicture(0, 0, 100, 100, imagePath);
        }

        // Save the workbook with all linked pictures added
        workbook.Save("BatchLinkedPictures.xlsx");
    }
}
