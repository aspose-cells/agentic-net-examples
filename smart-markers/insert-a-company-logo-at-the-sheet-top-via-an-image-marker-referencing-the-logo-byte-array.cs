using System;
using System.IO;
using Aspose.Cells;

class InsertLogoHeader
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Path to the company logo
            string logoPath = "company_logo.png";

            // Load the logo if the file exists
            if (File.Exists(logoPath))
            {
                byte[] logoBytes = File.ReadAllBytes(logoPath);

                // Insert the logo into the center section of the header (section index 1)
                // SetHeaderPicture adds the image data, SetHeader with "&G" displays the picture
                worksheet.PageSetup.SetHeaderPicture(1, logoBytes);
                worksheet.PageSetup.SetHeader(1, "&G");
            }
            else
            {
                Console.WriteLine($"Logo file not found: {logoPath}. Header will be saved without a picture.");
            }

            // Save the workbook with the header logo (or without if logo missing)
            string outputPath = "Workbook_With_Logo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}