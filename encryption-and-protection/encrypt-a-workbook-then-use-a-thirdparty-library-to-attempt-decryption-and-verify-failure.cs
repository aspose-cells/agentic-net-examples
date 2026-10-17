// Title: Encrypt an Excel workbook with a password using Aspose.Cells for .NET and confirm decryption fails when opened without the password
// Common Searches: how to set a password on an Excel workbook using Aspose.Cells C# | Aspose.Cells .NET open password‑protected workbook without providing a password throws exception | C# verify that an encrypted .xlsx cannot be opened without the password using Aspose.Cells | sample code for encrypting an Excel file with Aspose.Cells and handling the decryption error | Workbook.Settings.Password example for password protection in C#
// Tags: Aspose.Cells password protection API | C# encrypt Excel workbook Aspose.Cells | exception handling for protected workbook Aspose.Cells | saving encrypted .xlsx with Aspose.Cells | Workbook.Settings.Password property usage

using System;
using System.IO;
using Aspose.Cells;

// The program creates a simple workbook, assigns a password via Workbook.Settings.Password, saves it as an encrypted .xlsx file, then attempts to load the file without supplying the password. The expected exception is caught, confirming that decryption fails for password‑protected workbooks.
class Program
{
    static void Main()
    {
        try
        {
            // Create a simple workbook using Aspose.Cells
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Cells["A1"].PutValue("Hello Aspose!");

            // Encrypt the workbook with a password
            workbook.Settings.Password = "SecretPassword";

            // Save the encrypted workbook
            string encryptedPath = "encrypted.xlsx";
            workbook.Save(encryptedPath, SaveFormat.Xlsx);
            Console.WriteLine($"Encrypted workbook saved to '{encryptedPath}'.");

            // Attempt to open the encrypted file without providing a password
            // This should throw an exception because the file is protected.
            if (File.Exists(encryptedPath))
            {
                try
                {
                    // Aspose.Cells will raise an exception for password‑protected files
                    Workbook protectedWorkbook = new Workbook(encryptedPath);
                    // If no exception, decryption unexpectedly succeeded
                    Console.WriteLine("Unexpected success: Aspose.Cells opened the encrypted workbook without a password.");
                }
                catch (Exception ex)
                {
                    // Expected outcome: decryption fails
                    Console.WriteLine("Decryption failed as expected. Exception message:");
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine($"File not found: {encryptedPath}");
            }
        }
        catch (Exception ex)
        {
            // General safety catch‑all
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine(ex.Message);
        }
    }
}
