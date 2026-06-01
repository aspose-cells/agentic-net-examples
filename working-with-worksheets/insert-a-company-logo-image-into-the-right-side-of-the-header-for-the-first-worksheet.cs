using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsHeaderLogo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Define file paths
                string logoPath = "CompanyLogo.png";
                string outputPath = "WorkbookWithHeaderLogo.xlsx";

                // Verify the logo image exists
                if (!File.Exists(logoPath))
                {
                    Console.WriteLine($"Logo file not found: {logoPath}");
                    return;
                }

                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet worksheet = workbook.Worksheets[0];

                // Load the logo image into a byte array
                byte[] logoBytes = File.ReadAllBytes(logoPath);

                // Insert the image into the right side of the header (section index 2)
                worksheet.PageSetup.SetHeaderPicture(2, logoBytes);

                // Set the header script for the right section to display the picture
                worksheet.PageSetup.SetHeader(2, "&G");

                // Ensure the output directory exists
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}