// Title: C# – Update a Picture’s Linked Cell and Refresh the Image with Aspose.Cells
// Description: Demonstrates how to add a picture to a worksheet, link it to a cell, change the cell value, and call Picture.UpdateSelectedValue() to refresh the image before saving the workbook using Aspose.Cells for .NET.
// Keywords: Aspose.Cells picture linked cell | UpdateSelectedValue C# | refresh picture after cell change | Aspose.Cells .NET picture shape | link image to cell Aspose | C# workbook picture refresh | Aspose.Cells example
// Common Searches: Aspose.Cells how to refresh a picture after updating linked cell | C# set linked cell for picture shape Aspose.Cells | Update picture linked to cell programmatically | Picture.UpdateSelectedValue example .NET | link image to cell and refresh in Aspose.Cells
// Developer Intent: Refresh a picture shape after modifying the value of its linked cell in a .NET workbook.
// Use Cases: Create a dynamic report where a logo image reflects a title stored in a cell and updates automatically when the title changes. | Display status icons that change based on cell values; after recalculating the status, update the cells and refresh the icons. | Generate a batch of charts or symbols, each linked to different cells, then modify those cells with calculated data and refresh all pictures in one pass.
// AI Prompts: Generate C# code that links a picture to a specific cell using Aspose.Cells and refreshes the picture after the cell value is changed. | Show how to loop through multiple picture shapes, assign each a distinct linked cell, update the cells, and call UpdateSelectedValue for every picture. | Explain error handling for missing image files when adding pictures and ensure the picture is refreshed after cell updates in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to add a picture to a worksheet, link it to a cell, change the cell value, and call Picture.UpdateSelectedValue() to refresh the image before saving the workbook using Aspose.Cells for .NET.
class UpdatePictureLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Put an initial value into cell A1 (this will be linked to the picture)
            worksheet.Cells["A1"].PutValue("Initial Value");

            // Path to the picture file
            string picturePath = "sample.png";

            // Verify that the picture file exists before adding it
            if (!File.Exists(picturePath))
                throw new FileNotFoundException($"Picture file not found: {picturePath}");

            // Add a picture to the worksheet at cell C3 (row index 2, column index 2)
            int pictureIndex = worksheet.Pictures.Add(2, 2, picturePath);
            Picture picture = worksheet.Pictures[pictureIndex];

            // Link the picture to cell A1
            // Parameters: linked cell address, isRowAbsolute, isColumnAbsolute
            picture.SetLinkedCell("A1", false, false);

            // Change the linked cell's value; the picture will reflect this after refresh
            worksheet.Cells["A1"].PutValue("Updated Value");

            // Refresh the picture so it reflects the new linked cell value
            picture.UpdateSelectedValue();

            // Save the workbook
            string outputPath = "UpdatedPictureLinkedCell.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
