// Title: C# error handling to distinguish wrong password from corrupted encrypted Excel workbook with Aspose.Cells
// Common Searches: how to catch wrong password error when opening an encrypted .xlsx using Aspose.Cells in C# | detect corrupted encrypted Excel file versus invalid password with Aspose.Cells LoadOptions | Aspose.Cells CellsException message check for password mismatch in .NET | C# load encrypted workbook and differentiate password mismatch from file corruption | handle encrypted Excel file errors and provide specific messages with Aspose.Cells
// Tags: Aspose.Cells load encrypted workbook error handling | CellsException password mismatch detection | encrypted .xlsx file corruption handling Aspose.Cells | C# differentiate password error from file corruption | LoadOptions password validation Aspose.Cells

using System;
using Aspose.Cells;

// The example loads an encrypted Excel workbook using Aspose.Cells LoadOptions with a user‑supplied password, then catches CellsException to examine its message. If the message contains the word "password" a specific wrong‑password error is reported; otherwise a corrupted or unreadable file error is shown, with a generic catch for any other exceptions.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // The password supplied by the user (could be wrong)
        string suppliedPassword = "UserSuppliedPassword";

        // Prepare load options with the supplied password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        loadOptions.Password = suppliedPassword;

        try
        {
            // Attempt to load the encrypted workbook
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If we reach this point the password is correct and the file is not corrupted
            Console.WriteLine("Workbook loaded successfully.");
            // Further processing can be done here...
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for both incorrect password and corrupted file scenarios.
            // Distinguish based on the exception message or error code.

            // Incorrect password typically contains this phrase in the message.
            if (ex.Message != null && ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: The supplied password is incorrect.");
            }
            // Corrupted or unreadable file will have a different message.
            else
            {
                Console.WriteLine("Error: The encrypted file appears to be corrupted or unreadable.");
            }
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions.
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
