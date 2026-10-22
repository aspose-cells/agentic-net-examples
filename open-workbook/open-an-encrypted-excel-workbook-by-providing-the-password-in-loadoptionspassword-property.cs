// Title: Load a password‑protected Excel (.xlsx) workbook in C# using Aspose.Cells LoadOptions.Password
// Common Searches: C# Aspose.Cells load encrypted .xlsx file with password | How to open a password protected Excel workbook using LoadOptions in Aspose.Cells | Aspose.Cells Workbook constructor with password for encrypted workbook | Read encrypted Excel file in .NET by setting LoadOptions.Password | Decrypt an Excel workbook while loading with Aspose.Cells C# example
// Tags: Aspose.Cells load encrypted xlsx | LoadOptions password property | Workbook constructor password protected file | handle CellsException invalid password | save decrypted workbook Aspose.Cells

using Aspose.Cells;
using System;
using System.IO;

// The sample verifies the encrypted Excel file exists, creates a LoadOptions object with LoadFormat.Xlsx and the required password, loads the workbook via the password‑aware Workbook constructor, accesses the first worksheet, saves a decrypted copy, and catches CellsException to report an invalid password while handling other exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string encryptedFilePath = "EncryptedWorkbook.xlsx";

        // Verify that the source file exists
        if (!File.Exists(encryptedFilePath))
        {
            Console.WriteLine($"Error: File \"{encryptedFilePath}\" not found.");
            return;
        }

        try
        {
            // Set load options with the required password
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "YourPasswordHere"
            };

            // Load the workbook using the password‑protected options
            Workbook workbook = new Workbook(encryptedFilePath, loadOptions);

            // Access the first worksheet as an example operation
            Worksheet firstSheet = workbook.Worksheets[0];
            Console.WriteLine($"Successfully opened worksheet: {firstSheet.Name}");

            // Save the workbook to a new file after opening
            string decryptedFilePath = "DecryptedWorkbook.xlsx";

            // Ensure the target directory exists
            string outputDir = Path.GetDirectoryName(decryptedFilePath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            workbook.Save(decryptedFilePath);
            Console.WriteLine($"Workbook saved as \"{decryptedFilePath}\".");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for password issues; check the message
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The provided password is invalid.");
            }
            else
            {
                Console.WriteLine($"CellsException: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
