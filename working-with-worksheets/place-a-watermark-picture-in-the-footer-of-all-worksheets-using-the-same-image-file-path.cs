// Title: Add a uniform footer image watermark to every worksheet using Aspose.Cells for .NET (C#)
// Description: The sample reads a PNG file, creates or opens a Workbook, and inserts the image as a centered footer picture on each worksheet with PageSetup.SetFooterPicture and PageSetup.SetFooter("&G"). The file is saved as WorkbookWithFooterWatermark.xlsx.
// Keywords: Aspose.Cells C# footer image | Excel footer watermark .NET | SetFooterPicture Aspose | centered footer picture | same image all worksheets | Excel workbook watermark | read PNG bytes C# | PageSetup footer picture | Aspose.Cells tutorial | Excel automation watermark
// Common Searches: asp.net add image to Excel footer | c# Aspose.Cells set footer picture for all sheets | how to place a logo in Excel footer using Aspose | apply identical footer watermark across worksheets programmatically | set footer picture center section Aspose.Cells
// Developer Intent: Insert the same picture as a footer watermark on every worksheet in an Excel workbook using Aspose.Cells for .NET.
// Use Cases: Generate a report workbook that automatically displays the company logo in the footer of each sheet. | Add a confidentiality‑notice image to the footer of all worksheets in a financial statement before distribution. | Create monthly sales workbooks where a branding graphic appears consistently in the footer of every page.
// AI Prompts: Show how to move the watermark to the left or right footer section instead of the center. | Provide an example that reads the image from a MemoryStream and applies it to all worksheets. | Explain techniques for resizing or compressing a large PNG before calling SetFooterPicture to improve performance.

using System;
using System.IO;
using Aspose.Cells;

// The sample reads a PNG file, creates or opens a Workbook, and inserts the image as a centered footer picture on each worksheet with PageSetup.SetFooterPicture and PageSetup.SetFooter("&G"). The file is saved as WorkbookWithFooterWatermark.xlsx.
class Program
{
    static void Main()
    {
        // Path to the watermark image file
        string imagePath = "watermark.png";

        byte[] imageData = null;

        // Verify that the image file exists before attempting to read it
        if (File.Exists(imagePath))
        {
            try
            {
                imageData = File.ReadAllBytes(imagePath);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading image file: {ex.Message}");
            }
        }
        else
        {
            Console.WriteLine($"Image file not found: {imagePath}. Watermark will be skipped.");
        }

        // Create a new workbook (or load an existing one)
        Workbook workbook = new Workbook();

        // Apply the footer picture to every worksheet if image data is available
        if (imageData != null)
        {
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Set the picture in the center section of the footer (section index 1)
                    sheet.PageSetup.SetFooterPicture(1, imageData);

                    // Use the image placeholder "&G" to display the picture in the footer
                    sheet.PageSetup.SetFooter(1, "&G");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error applying watermark to sheet '{sheet.Name}': {ex.Message}");
                }
            }
        }

        // Save the modified workbook
        try
        {
            workbook.Save("WorkbookWithFooterWatermark.xlsx");
            Console.WriteLine("Workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error saving workbook: {ex.Message}");
        }
    }
}
