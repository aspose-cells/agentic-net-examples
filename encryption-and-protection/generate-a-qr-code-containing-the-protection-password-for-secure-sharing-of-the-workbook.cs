// Title: Generate a QR code for the workbook protection password and embed it in an Excel worksheet using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# create QR code image from password string | how to protect an Excel workbook and add a QR code with Aspose.Cells | Worksheet.Pictures.Add QR code example in .NET | set custom size for QR code picture in Aspose.Cells workbook
// Tags: Aspose.Cells QR code generation from string | Workbook.Protect password protection C# | Worksheet.Pictures.Add embedding QR code | adjust QR code image dimensions Aspose.Cells | save protected workbook as Xlsx Aspose.Cells

using System;
using Aspose.Cells;
using Aspose.Cells.Drawing;

// The sample creates a new Workbook, applies password protection with Workbook.Protect, generates a QR code picture that encodes the password using Worksheet.Pictures.Add, customizes the QR code image size, and saves the protected workbook as an Xlsx file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Define the protection password
            string password = "SecretPassword123";

            // Protect the entire workbook with the password
            workbook.Protect(ProtectionType.All, password);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Add a QR code picture that contains the password.
            // The overload that takes a string creates a QR code automatically.
            int pictureIndex = sheet.Pictures.Add(0, 0, password);
            Picture qrPicture = sheet.Pictures[pictureIndex];
            qrPicture.Width = 150;   // Set desired width
            qrPicture.Height = 150;  // Set desired height

            // Save the workbook to a file
            string outputPath = "ProtectedWorkbook.xlsx";
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
