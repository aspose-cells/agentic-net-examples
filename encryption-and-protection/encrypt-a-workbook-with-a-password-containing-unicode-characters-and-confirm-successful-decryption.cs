// Title: Encrypt an Excel workbook with a Unicode password and validate decryption using Aspose.Cells for .NET
// AI Prompts: Set wb.Settings.Password to a Unicode string (including accented characters or emojis) and save the workbook as an encrypted .xlsx file with Aspose.Cells. | Load the encrypted .xlsx file using LoadOptions.Password set to the same Unicode password, then read a cell to confirm the original content is intact. | Replace the Unicode password with a different non‑ASCII string, re‑encrypt the workbook, and repeat the load‑and‑verify steps to demonstrate consistent handling.
// Common Searches: asp.net set Unicode password for Excel file encryption with Aspose.Cells | load password protected xlsx containing emoji using Aspose.Cells C# | verify workbook decryption after applying non‑ASCII password in Aspose.Cells | how to encrypt Excel workbook with special characters in password using Aspose.Cells .NET
// Tags: unicode password encryption Aspose.Cells | load encrypted workbook with LoadOptions password | c# verify workbook decryption Aspose.Cells | excel file protection non-ascii password .NET | save encrypted xlsx using wb.Settings.Password

using System;
using Aspose.Cells;

// The example creates a workbook, assigns a Unicode password (e.g., containing accented letters and an emoji) via wb.Settings.Password, saves it as an encrypted .xlsx file, then reloads it with LoadOptions.Password and checks that cell A1 still contains "Hello, World!" to confirm successful decryption.
class Program
{
    static void Main()
    {
        // Define a Unicode password
        string unicodePassword = "Pässwörd🔐";

        // -------------------------
        // Create a new workbook
        // -------------------------
        Workbook wb = new Workbook();

        // Add some sample data to the first worksheet
        Worksheet sheet = wb.Worksheets[0];
        sheet.Cells["A1"].PutValue("Hello, World!");
        sheet.Cells["A2"].PutValue(DateTime.Now);

        // Set the password for encryption (Unicode supported)
        wb.Settings.Password = unicodePassword;

        // Save the encrypted workbook
        string encryptedPath = "encrypted.xlsx";
        wb.Save(encryptedPath);

        // -------------------------
        // Load the workbook with the password
        // -------------------------
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = unicodePassword;

        Workbook loadedWb = new Workbook(encryptedPath, loadOptions);

        // Verify that the data was decrypted correctly
        Worksheet loadedSheet = loadedWb.Worksheets[0];
        string cellValue = loadedSheet.Cells["A1"].StringValue;

        if (cellValue == "Hello, World!")
        {
            Console.WriteLine("Workbook decrypted successfully with Unicode password.");
        }
        else
        {
            Console.WriteLine("Decryption failed or data corrupted.");
        }
    }
}
