// Title: Insert a company logo into the top of an Excel worksheet using Aspose.Cells Pictures.Add with a byte array in C#
// Common Searches: Aspose.Cells C# insert logo from byte array into Excel worksheet | how to add an image to cell A1 using MemoryStream in Aspose.Cells | C# programmatically place a PNG logo at the top of an Excel file with Aspose | using Worksheet.Pictures.Add to embed a company logo in .xlsx via byte[] | Aspose.Cells picture insertion from file bytes example .NET
// Tags: Worksheet.Pictures.Add byte array | Aspose.Cells insert image from MemoryStream | C# add logo to Excel top row | programmatic Excel logo insertion .NET | embed PNG logo in workbook using Aspose.Cells | Excel picture placement at cell A1 C#

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a new workbook, reads a PNG logo into a byte array, wraps it in a MemoryStream, and uses Worksheet.Pictures.Add to place the image at cell A1 (the sheet top‑left). It then saves the workbook as OutputWithLogo.xlsx, handling missing output directories and missing logo files.
class InsertLogoExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Load the logo image into a byte array
            byte[] logoBytes = GetCompanyLogoBytes();

            // Insert the logo at cell A1 using a MemoryStream (required overload)
            using (MemoryStream ms = new MemoryStream(logoBytes))
            {
                // upperLeftRow, upperLeftColumn, lowerRightRow, lowerRightColumn, pictureStream
                sheet.Pictures.Add(0, 0, 0, 0, ms);
            }

            // Ensure the output directory exists (handle cases where outputPath has no directory)
            string outputPath = "OutputWithLogo.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }

    // Retrieves the logo bytes from a file, ensuring the file exists.
    static byte[] GetCompanyLogoBytes()
    {
        string logoPath = "company_logo.png";

        if (!File.Exists(logoPath))
        {
            throw new FileNotFoundException($"Logo file not found: {logoPath}");
        }

        return File.ReadAllBytes(logoPath);
    }
}
