// Title: Validate that an Excel workbook encrypted with a password cannot be opened using a different similar password in C# with Aspose.Cells
// Common Searches: how to verify that an encrypted Excel file rejects an incorrect password using Aspose.Cells C# | Aspose.Cells C# load encrypted workbook with wrong password throws exception | test workbook password protection with similar but wrong password in .NET | catch CellsException when opening password‑protected Excel with mismatched password
// Tags: encrypt workbook with password Aspose.Cells | open password‑protected Excel with wrong password C# | CellsException handling for password mismatch | LoadOptions password validation Aspose.Cells | validate workbook encryption integrity .NET

using System;
using Aspose.Cells;
using System.IO;

// The example creates a workbook, applies a password to encrypt it, saves it as an .xlsx file, then attempts to open the file using LoadOptions with a similar but incorrect password. Aspose.Cells throws a CellsException, which is caught to confirm that the wrong password is rejected, and the temporary file is cleaned up.
class Program
{
    static void Main()
    {
        // Define passwords
        string correctPassword = "Secret123!";
        string wrongPassword = "Secret124!"; // similar but different

        // Create a temporary file path
        string filePath = Path.Combine(Path.GetTempPath(), "EncryptedWorkbook.xlsx");

        // ------------------- Create and encrypt workbook -------------------
        // Create a new workbook with a single worksheet
        Workbook wb = new Workbook();
        Worksheet ws = wb.Worksheets[0];
        ws.Cells["A1"].PutValue("Hello, Aspose!");

        // Set the password for encryption
        wb.Settings.Password = correctPassword;

        // Save the encrypted workbook
        wb.Save(filePath);
        Console.WriteLine($"Workbook saved and encrypted at: {filePath}");

        // ------------------- Attempt to open with wrong password -------------------
        try
        {
            // LoadOptions allows specifying a password for opening
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = wrongPassword
            };

            // Attempt to load the encrypted workbook with the wrong password
            Workbook wbWrong = new Workbook(filePath, loadOptions);

            // If no exception, the workbook opened (unexpected)
            Console.WriteLine("Unexpectedly opened the workbook with the wrong password.");
        }
        catch (CellsException ex)
        {
            // Expected path: Aspose.Cells throws an exception when password is incorrect
            Console.WriteLine("Failed to open workbook with wrong password as expected.");
            Console.WriteLine($"Exception message: {ex.Message}");
        }
        finally
        {
            // Clean up the temporary file
            if (File.Exists(filePath))
            {
                File.Delete(filePath);
                Console.WriteLine("Temporary file deleted.");
            }
        }
    }
}
