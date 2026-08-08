// Title: Add a Company Logo to an Excel Header via Byte Array with Aspose.Cells for .NET
// Description: Creates a new workbook, loads a PNG logo into a byte array, inserts the image into the center header using Worksheet.PageSetup.SetHeaderPicture and the &G placeholder, optionally scales the picture, and saves the file.
// Keywords: Aspose.Cells header picture | C# insert logo Excel header | byte array image Excel | SetHeaderPicture Aspose | Excel &G placeholder | scale header image Aspose.Cells
// Common Searches: how to add logo to Excel header Aspose.Cells | set header picture from byte array C# | use &G placeholder for header image | resize header logo Aspose.Cells | insert image into Excel page header programmatically
// Developer Intent: Place a company logo in the top header of an Excel worksheet by loading the image as a byte array.
// Use Cases: Automated generation of branded reports with a centered logo on every printed page. | Creating invoices that embed the corporate logo in the header without manual editing. | Producing printable worksheets where the logo size adapts to page layout.
// AI Prompts: Show how to set a header picture from a MemoryStream instead of a file using Aspose.Cells for .NET. | Generate code that calculates dynamic scaling for a header logo based on page width and height. | Explain how to assign different images to the left, center, and right header sections in Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// Creates a new workbook, loads a PNG logo into a byte array, inserts the image into the center header using Worksheet.PageSetup.SetHeaderPicture and the &G placeholder, optionally scales the picture, and saves the file.
class InsertLogoInHeader
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            string logoPath = "company_logo.png";

            if (File.Exists(logoPath))
            {
                // Load the company logo into a byte array
                byte[] logoBytes = File.ReadAllBytes(logoPath);

                // Insert the logo into the center section of the header (section index 1)
                Picture headerPic = worksheet.PageSetup.SetHeaderPicture(1, logoBytes);

                // Set the header script to display the picture placeholder (&G)
                worksheet.PageSetup.SetHeader(1, "&G");

                // Optional: adjust picture scaling
                // headerPic.ScaleX = 0.5; // 50% width
                // headerPic.ScaleY = 0.5; // 50% height
            }
            else
            {
                Console.WriteLine($"Logo file not found: {logoPath}. Header will be created without image.");
                // Optionally set a text header instead
                worksheet.PageSetup.SetHeader(1, "Company Header");
            }

            // Save the workbook
            string outputPath = "Workbook_With_Logo_Header.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
