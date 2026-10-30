// Title: Apply AES‑256 encryption to an Excel workbook with Aspose.Cells for .NET and compare original and encrypted file sizes
// Common Searches: how to encrypt an XLSX file with AES 256 using Aspose.Cells in C# | retrieve file size of a password‑protected Excel workbook with Aspose.Cells .NET | compare original and encrypted Excel file sizes programmatically in C# | set workbook password for AES‑256 encryption using Aspose.Cells API | measure impact of encryption on XLSX file size with Aspose.Cells
// Tags: AES-256 encryption with Aspose.Cells | set workbook password Aspose.Cells .NET | XLSX file size comparison after encryption | C# retrieve saved workbook size | Aspose.Cells encryption impact on file size

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, adds sample data, saves it as an unencrypted XLSX file, records its size, applies a strong password that triggers AES‑256 encryption via the Workbook.Settings.Password property, saves the encrypted file, records the new size, and prints both sizes for easy comparison, with error handling for robustness.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add sample data to the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Sample Text");
            sheet.Cells["B1"].PutValue(12345);

            // Save the original (unencrypted) workbook
            string originalPath = "original.xlsx";
            workbook.Save(originalPath, SaveFormat.Xlsx);

            // Get size of the original file if it exists
            long originalSize = File.Exists(originalPath) ? new FileInfo(originalPath).Length : 0;

            // Apply strong encryption (AES 256) with a password
            workbook.Settings.Password = "StrongPassword!@#123";

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);

            // Get size of the encrypted file if it exists
            long encryptedSize = File.Exists(encryptedPath) ? new FileInfo(encryptedPath).Length : 0;

            // Output the file sizes for comparison
            Console.WriteLine($"Original file size: {originalSize} bytes");
            Console.WriteLine($"Encrypted file size: {encryptedSize} bytes");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
