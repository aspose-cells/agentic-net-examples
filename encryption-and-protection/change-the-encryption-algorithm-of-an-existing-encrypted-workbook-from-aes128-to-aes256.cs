// Title: Re‑encrypt an Excel workbook from AES‑128 to AES‑256 with Aspose.Cells for .NET (C#)
// Description: Loads a password‑protected .xlsx encrypted with AES‑128, keeps the original password, switches the encryption to AES‑256 using Workbook.SetEncryptionOptions(StrongCryptographicProvider, 256), and saves the file. Includes robust error handling for missing files and Aspose.Cells exceptions.
// Keywords: Aspose.Cells | C# | AES-128 | AES-256 | encryption upgrade | Workbook.SetEncryptionOptions | StrongCryptographicProvider | re‑encrypt Excel | password‑protected workbook | SaveFormat.Xlsx | LoadOptions
// Common Searches: change Excel encryption from AES-128 to AES-256 Aspose.Cells | C# re‑encrypt password protected workbook Aspose.Cells | Aspose.Cells set encryption to 256‑bit | upgrade Excel file security Aspose.Cells .NET | how to use StrongCryptographicProvider with Aspose.Cells
// Developer Intent: Upgrade an existing AES‑128 encrypted workbook to AES‑256 while preserving its password.
// Use Cases: Align legacy spreadsheets with modern security standards. | Strengthen protection before distributing files to partners. | Automate batch conversion of multiple AES‑128 workbooks to AES‑256.
// AI Prompts: Show C# code that loads an AES‑128 encrypted .xlsx with Aspose.Cells, changes the encryption to AES‑256, and saves it using the same password. | Explain the parameters of Workbook.SetEncryptionOptions for enabling AES‑256 in Aspose.Cells. | Create a try‑catch structure that captures CellsException when re‑encrypting a workbook with StrongCryptographicProvider.

using System;
using System.IO;
using Aspose.Cells;

// Loads a password‑protected .xlsx encrypted with AES‑128, keeps the original password, switches the encryption to AES‑256 using Workbook.SetEncryptionOptions(StrongCryptographicProvider, 256), and saves the file. Includes robust error handling for missing files and Aspose.Cells exceptions.
class ChangeEncryption
{
    static void Main()
    {
        // Paths to the original and the re‑encrypted workbook
        string sourcePath = "encrypted128.xlsx";
        string destPath = "encrypted256.xlsx";

        // Password used for the original encryption
        string password = "myPassword";

        try
        {
            // Verify that the source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the existing encrypted workbook using the password
            LoadOptions loadOptions = new LoadOptions
            {
                Password = password
            };
            Workbook workbook = new Workbook(sourcePath, loadOptions);

            // Preserve the password for saving
            workbook.Settings.Password = password;

            // Change encryption to AES‑256 (StrongCryptographicProvider with 256‑bit key)
            workbook.SetEncryptionOptions(EncryptionType.StrongCryptographicProvider, 256);

            // Save the workbook with the new encryption settings
            workbook.Save(destPath, SaveFormat.Xlsx);
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine($"File error: {ex.Message}");
        }
        catch (CellsException ex)
        {
            Console.WriteLine($"Aspose.Cells error: {ex.Message}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
