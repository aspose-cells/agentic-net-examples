// Title: Retry opening an encrypted Excel workbook with an alternate password using Aspose.Cells in C#
// Common Searches: c# Aspose.Cells load encrypted workbook with secondary password fallback | how to handle password failure when opening a password‑protected Excel file using Aspose.Cells | retry loading encrypted .xlsx with another password Aspose.Cells .NET example | Aspose.Cells exception handling for wrong password and retry with another password | load encrypted Excel file using LoadOptions password alternative in C#
// Tags: Aspose.Cells fallback password loading | load encrypted workbook with alternate password .NET | exception handling workbook decryption Aspose.Cells | retry LoadOptions password C# | decrypt Excel file and save copy Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample checks for an encrypted Excel file, attempts to open it with a primary password via LoadOptions, and if that fails, catches the exception and retries using an alternate password. Successful loading allows the workbook to be saved as a decrypted copy, while failures are logged with both error messages.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Primary and alternate passwords
        string primaryPassword = "FirstPassword";
        string alternatePassword = "SecondPassword";

        // Verify that the source file exists
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found – {filePath}");
            return;
        }

        Workbook workbook = null;

        // Try loading with the primary password
        try
        {
            var loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = primaryPassword
            };
            workbook = new Workbook(filePath, loadOptions);
        }
        catch (Exception exPrimary)
        {
            // If loading fails, attempt with the alternate password
            try
            {
                var loadOptions = new LoadOptions(LoadFormat.Xlsx)
                {
                    Password = alternatePassword
                };
                workbook = new Workbook(filePath, loadOptions);
            }
            catch (Exception exAlternate)
            {
                Console.WriteLine("Failed to open the workbook with both passwords.");
                Console.WriteLine($"Primary attempt error: {exPrimary.Message}");
                Console.WriteLine($"Alternate attempt error: {exAlternate.Message}");
                return;
            }
        }

        // At this point the workbook is loaded successfully
        // Example: save a decrypted copy of the workbook
        try
        {
            workbook.Save("decrypted_copy.xlsx");
            Console.WriteLine("Decrypted copy saved as 'decrypted_copy.xlsx'.");
        }
        catch (Exception saveEx)
        {
            Console.WriteLine($"Error saving decrypted workbook: {saveEx.Message}");
        }
    }
}
