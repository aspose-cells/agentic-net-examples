// Title: Validate that a password‑protected Excel workbook cannot be opened with an incorrect password using Aspose.Cells for .NET
// AI Prompts: Show how to configure Aspose.Cells to encrypt a workbook with AES‑256, save it as .xlsx, and then try to open it using an AES‑128 password, capturing the resulting CellsException. | Generate C# code that creates a workbook, applies password protection, saves it, and demonstrates error handling when loading the file with a wrong password. | Provide a step‑by‑step example of using LoadOptions with mismatched encryption to confirm that Aspose.Cells throws a CellsException for a password mismatch.
// Common Searches: aspnet aspose.cells verify workbook cannot be opened with incorrect password | c# aspose.cells aes-256 encryption test with aes-128 password mismatch | how to catch CellsException when loading password protected xlsx with wrong password | aspose.cells loadoptions incorrect password example | validate encryption algorithm mismatch in excel file using aspose.cells
// Tags: AES-256 encryption Aspose.Cells .NET | incorrect password handling Aspose.Cells | CellsException password mismatch detection | load password‑protected XLSX with LoadOptions | Excel workbook protection validation Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample creates a Workbook, sets a password (default encryption), saves it as an .xlsx file, then attempts to open the same file using LoadOptions with an incorrect password. The operation catches a CellsException, confirming that the workbook cannot be opened when the provided password does not match the one used during encryption.
class Program
{
    static void Main()
    {
        try
        {
            // Create a workbook and protect it with a password (default encryption)
            Workbook wb = new Workbook();
            wb.Worksheets[0].Cells["A1"].PutValue("Sensitive Data");
            wb.Settings.Password = "StrongPassword123!";

            string filePath = "EncryptedWorkbook.xlsx";

            // Save the protected workbook
            wb.Save(filePath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved to '{filePath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error during workbook creation/saving: {ex.Message}");
            return;
        }

        // Attempt to open the protected workbook using an incorrect password
        try
        {
            string filePath = "EncryptedWorkbook.xlsx";

            if (!File.Exists(filePath))
                throw new FileNotFoundException("The encrypted workbook file was not found.", filePath);

            LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
            {
                Password = "WrongPassword"
            };

            Workbook loadedWb = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened unexpectedly with an incorrect password.");
        }
        catch (CellsException ex)
        {
            // Expected: the password is incorrect
            Console.WriteLine("Failed to open workbook with incorrect password as expected.");
            Console.WriteLine("Exception message: " + ex.Message);
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine("File not found: " + ex.Message);
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions
            Console.WriteLine("An unexpected error occurred: " + ex.Message);
        }
    }
}
