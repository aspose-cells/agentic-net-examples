// Title: Refresh a Picture Shape Linked to a Cell with Aspose.Cells for .NET (C#)
// Description: Shows how to generate placeholder PNG files, insert a picture on a worksheet, link the picture to a cell, modify the cell to point to a new image file, invoke UpdateSelectedValue to refresh the picture, and save the workbook.
// Keywords: Aspose.Cells | C# | .NET | picture shape | linked cell | UpdateSelectedValue | refresh image | dynamic picture | Excel automation | SetLinkedCell
// Common Searches: Aspose.Cells change picture source via linked cell | UpdateSelectedValue example C# | link Excel picture to cell and refresh | replace image in picture shape programmatically | how to refresh linked picture in Aspose.Cells
// Developer Intent: Refresh a picture shape after the value of its linked cell is updated.
// Use Cases: Swap product photos in a generated report by storing file paths in cells. | Allow end‑users to edit image file names in a template and have the workbook automatically display the new images. | Create invoices where the company logo can be changed through a cell without recreating the picture object.
// AI Prompts: Provide a C# example that links a picture to a cell and updates the picture after changing the cell value using Aspose.Cells. | Explain the purpose of picture.UpdateSelectedValue and the required format of the linked cell content. | Show step‑by‑step code to create placeholder PNG files, add a picture, set a linked cell, modify the cell, and refresh the image.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Shows how to generate placeholder PNG files, insert a picture on a worksheet, link the picture to a cell, modify the cell to point to a new image file, invoke UpdateSelectedValue to refresh the picture, and save the workbook.
class UpdatePictureLinkedCellExample
{
    static void Main()
    {
        try
        {
            // Minimal 1x1 PNG (transparent) used as placeholder image
            byte[] pngData = new byte[]
            {
                0x89,0x50,0x4E,0x47,0x0D,0x0A,0x1A,0x0A,
                0x00,0x00,0x00,0x0D,0x49,0x48,0x44,0x52,
                0x00,0x00,0x00,0x01,0x00,0x00,0x00,0x01,
                0x08,0x06,0x00,0x00,0x00,0x1F,0x15,0xC4,
                0x89,0x00,0x00,0x00,0x0A,0x49,0x44,0x41,
                0x54,0x78,0x9C,0x63,0x60,0x00,0x00,0x00,
                0x02,0x00,0x01,0xE2,0x21,0xBC,0x33,0x00,
                0x00,0x00,0x00,0x49,0x45,0x4E,0x44,0xAE,
                0x42,0x60,0x82
            };

            // Ensure placeholder image exists
            string placeholderPath = "Aspose.Cells.png";
            EnsureImageExists(placeholderPath, pngData);

            // Ensure second image exists for update demonstration (reuse same PNG)
            string secondImagePath = "Aspose.CellsLogo.png";
            EnsureImageExists(secondImagePath, pngData);

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Put initial image file name into linked cell B2
            sheet.Cells["B2"].PutValue(placeholderPath);

            // Add picture using the placeholder image
            int pictureIndex = sheet.Pictures.Add(5, 2, 100, 100, placeholderPath);
            Picture picture = sheet.Pictures[pictureIndex];

            // Link picture to cell B2
            picture.SetLinkedCell("B2", false, false);

            // Change linked cell to second image file name
            sheet.Cells["B2"].PutValue(secondImagePath);

            // Refresh picture based on linked cell value
            picture.UpdateSelectedValue();

            // Save the workbook
            workbook.Save("UpdatedPictureLinkedCell.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }

    // Creates a PNG file at the specified path if it does not already exist
    private static void EnsureImageExists(string path, byte[] pngBytes)
    {
        try
        {
            if (!File.Exists(path))
            {
                File.WriteAllBytes(path, pngBytes);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Failed to create image '{path}': {ex.Message}");
            throw;
        }
    }
}
