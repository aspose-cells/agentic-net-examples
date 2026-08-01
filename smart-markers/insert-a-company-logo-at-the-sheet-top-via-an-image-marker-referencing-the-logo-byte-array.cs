// Title: C# – Insert Company Logo into Excel Header with Aspose.Cells SetHeaderPicture
// Description: Demonstrates how to load a PNG logo into a byte array and embed it in the center section of an Excel worksheet header using Aspose.Cells. The example sets the header script, adjusts the header margin for proper fit, includes robust file‑existence checking, and saves the workbook as an XLSX file.
// Keywords: Aspose.Cells C# header image | SetHeaderPicture byte array | add logo to Excel header .NET | Excel header picture Aspose | C# workbook header logo example | Aspose.Cells SetHeader usage | Excel branding with header logo | smart markers header image
// Common Searches: Aspose.Cells insert logo into Excel header C# | SetHeaderPicture example with byte array | how to add a company logo to Excel header using Aspose | center header image Aspose.Cells .NET | adjust Excel header margin for logo Aspose
// Developer Intent: Embed a company logo in the worksheet header by loading the image as a byte array and using SetHeaderPicture.
// Use Cases: Generate branded reports where the logo appears on every printed page. | Automate invoice creation with a consistent header logo for corporate identity. | Produce Excel exports that include a centered header image and proper margin settings for printing.
// AI Prompts: Write C# code with Aspose.Cells that reads a PNG logo, sets it as the center header image using SetHeaderPicture, and adjusts the header margin. | Explain the relationship between SetHeaderPicture and SetHeader for embedding images in Excel headers with Aspose.Cells. | Provide error‑handling patterns for missing logo files and ensure the workbook saves after adding a header picture.

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates how to load a PNG logo into a byte array and embed it in the center section of an Excel worksheet header using Aspose.Cells. The example sets the header script, adjusts the header margin for proper fit, includes robust file‑existence checking, and saves the workbook as an XLSX file.
class InsertLogoInHeader
{
    static void Main()
    {
        try
        {
            // Verify that the logo file exists to avoid FileNotFoundException
            const string logoPath = "CompanyLogo.png";
            if (!File.Exists(logoPath))
                throw new FileNotFoundException($"Logo file not found: {logoPath}");

            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Load the company logo into a byte array
            byte[] logoBytes = File.ReadAllBytes(logoPath);

            // Insert the logo into the center section of the header (section index 1)
            sheet.PageSetup.SetHeaderPicture(1, logoBytes);
            // Place the picture in the center of the header using the appropriate script
            sheet.PageSetup.SetHeader(1, "&C&G");

            // Optionally adjust the header margin to ensure the logo fits nicely (in inches)
            sheet.PageSetup.HeaderMargin = 0.5;

            // Save the workbook
            const string outputPath = "WorkbookWithLogoHeader.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
