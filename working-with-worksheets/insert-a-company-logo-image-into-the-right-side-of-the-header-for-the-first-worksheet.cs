// Title: Add a Company Logo to the Right Header of the First Worksheet (C# Aspose.Cells)
// Description: Creates a new workbook, loads a PNG logo, inserts it on the right side of the header using Worksheet.PageSetup.SetHeaderPicture (section 2) and the "&G" placeholder, then saves the file as an .xlsx document.
// Keywords: Aspose.Cells C# header image | SetHeaderPicture example | insert logo into Excel header | right side header picture Aspose | C# add image to worksheet header | Excel branding with Aspose.Cells
// Common Searches: Aspose.Cells add logo to header C# | SetHeaderPicture right side Excel | how to place image in Excel header using Aspose | C# insert PNG into worksheet header | display picture in Excel header Aspose.Cells
// Developer Intent: Place a PNG company logo on the right side of the first worksheet's header and save the workbook.
// Use Cases: Generate branded reports with the corporate logo displayed on every printed page. | Automate invoice creation that includes a consistent header logo for company identity. | Produce marketing spreadsheets where the header reinforces visual branding.
// AI Prompts: Provide a C# snippet that adds a logo to the left side of the header with Aspose.Cells. | Explain how to set different header images for odd and even pages using SetHeaderPicture. | Show how to resize or scale a header picture inserted via SetHeaderPicture in C#.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, loads a PNG logo, inserts it on the right side of the header using Worksheet.PageSetup.SetHeaderPicture (section 2) and the "&G" placeholder, then saves the file as an .xlsx document.
class InsertHeaderLogo
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            string logoPath = "company_logo.png";

            // Verify that the logo file exists before attempting to read it
            if (File.Exists(logoPath))
            {
                // Load the company logo image file into a byte array
                byte[] logoBytes = File.ReadAllBytes(logoPath);

                // Insert the image into the right side of the header (section index 2)
                Picture headerPic = worksheet.PageSetup.SetHeaderPicture(2, logoBytes);

                // Set the header script to display the picture
                worksheet.PageSetup.SetHeader(2, "&G");
            }
            else
            {
                Console.WriteLine($"Logo file not found: {logoPath}. Header will be saved without image.");
            }

            // Save the workbook with the header image
            string outputPath = "WorkbookWithHeaderLogo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
