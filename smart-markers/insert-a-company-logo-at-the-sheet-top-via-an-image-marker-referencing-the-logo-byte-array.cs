// Title: Add a company PNG logo to the center header of an Excel workbook using Aspose.Cells SetHeaderPicture with a byte array (C#)
// AI Prompts: Load a PNG file into a byte array and use Aspose.Cells to place it in the header's center section. | Replace the header text with the &G placeholder and enable scaling for the inserted picture. | Add logic that sets a plain‑text header when the logo file cannot be found.
// Common Searches: C# Aspose.Cells how to set an image as the header picture using a byte array | Insert company logo into Excel header with Aspose.Cells SetHeaderPicture | Set center header image placeholder &G in Aspose.Cells .NET | Use SetHeaderPicture to add PNG logo to Excel page header in C# | Aspose.Cells header picture scaling options
// Tags: Aspose.Cells SetHeaderPicture byte array | header image insertion Aspose.Cells | header picture scaling C# | fallback plain text header Aspose.Cells | load PNG logo into Excel header

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeCellsLogoHeader
{
    // The example creates a new workbook, reads a PNG logo into a byte array, inserts it into the center of the page header with Worksheet.PageSetup.SetHeaderPicture, sets the header text to the &G placeholder, optionally enables picture scaling, and saves the file. If the logo file is missing, it falls back to a plain‑text header.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Path to the logo file
                string logoPath = "company_logo.png";

                // If the logo file exists, load it and insert into the header
                if (File.Exists(logoPath))
                {
                    byte[] logoBytes = File.ReadAllBytes(logoPath);

                    // Insert the logo into the center section of the header (1 = center)
                    Picture headerPicture = worksheet.PageSetup.SetHeaderPicture(1, logoBytes);

                    // Set the header text to display the picture placeholder (&G)
                    worksheet.PageSetup.SetHeader(1, "&G");

                    // Optional: adjust picture scaling
                    // headerPicture.ScaleToFit = true;
                }
                else
                {
                    // Logo file not found – optionally set a plain text header or leave it empty
                    worksheet.PageSetup.SetHeader(1, "Company Header");
                }

                // Save the workbook
                workbook.Save("Workbook_With_Logo_Header.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
