// Title: Insert Company Logo into Right Header of First Worksheet with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to load a PNG logo, use PageSetup.SetHeaderPicture to place it in the right section of the first worksheet's header, set the "&G" placeholder, and save the workbook as an .xlsx file.
// Keywords: Aspose.Cells C# header image | SetHeaderPicture right header | add logo to Excel header .NET | page setup header picture example | Aspose.Cells workbook branding | C# insert header picture | Excel header logo Aspose
// Common Searches: Aspose.Cells add logo to right header C# | SetHeaderPicture section 2 example | how to place an image in Excel header using Aspose | C# code for header picture in first worksheet | insert company logo into Excel header Aspose.Cells
// Developer Intent: Place a PNG logo in the right side of the header of the first worksheet and generate a branded Excel file.
// Use Cases: Generate corporate reports with a consistent logo on every printed page. | Automate invoice creation where the vendor's emblem appears in the header. | Create a reusable workbook template that includes branding in the header.
// AI Prompts: Write C# code with Aspose.Cells to add a PNG logo to the left header of all worksheets. | Show how to resize and align a header picture set via SetHeaderPicture in Aspose.Cells. | Provide robust error handling for missing image files when inserting a header picture with Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Demonstrates how to load a PNG logo, use PageSetup.SetHeaderPicture to place it in the right section of the first worksheet's header, set the "&G" placeholder, and save the workbook as an .xlsx file.
class InsertHeaderLogo
{
    static void Main()
    {
        try
        {
            // Verify that the logo image file exists
            const string logoPath = "logo.png";
            if (!File.Exists(logoPath))
            {
                Console.WriteLine($"Logo file not found: {logoPath}");
                return;
            }

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Load the company logo image into a byte array
            byte[] logoBytes = File.ReadAllBytes(logoPath);

            // Insert the image into the right section of the header (section index 2)
            // SetHeaderPicture returns a Picture object which can be further customized if needed
            Picture headerPic = worksheet.PageSetup.SetHeaderPicture(2, logoBytes);

            // Set the header script for the right section to display the picture
            // The "&G" placeholder tells Excel to render the picture set above
            worksheet.PageSetup.SetHeader(2, "&G");

            // Save the workbook with the header image
            const string outputPath = "Workbook_With_Header_Logo.xlsx";
            workbook.Save(outputPath);

            Console.WriteLine("Header logo inserted and workbook saved successfully.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
