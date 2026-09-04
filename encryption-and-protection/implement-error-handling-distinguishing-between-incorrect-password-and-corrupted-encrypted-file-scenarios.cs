// Title: Handle incorrect password and corrupted encrypted Excel files separately with Aspose.Cells in C#
// AI Prompts: Generate C# code that opens an encrypted .xlsx using Aspose.Cells, verifies the file exists, and prints a distinct message for an invalid password versus a corrupted workbook. | Write a try‑catch structure for Aspose.Cells Workbook loading that catches CellsException and uses the exception message to differentiate password errors from file corruption. | Show how to add a fallback catch for unexpected exceptions when loading a password‑protected Excel file with Aspose.Cells.
// Common Searches: Aspose.Cells C# detect wrong password when opening encrypted .xlsx | How to identify corrupted encrypted Excel file using Aspose.Cells .NET | C# differentiate CellsException password error from file corruption in Aspose.Cells | Load encrypted workbook with Aspose.Cells and handle invalid password separately | Aspose.Cells error handling for password‑protected Excel files in C#
// Tags: Aspose.Cells load encrypted xlsx with password | CellsException password error handling | encrypted workbook corruption detection Aspose.Cells | C# file existence check before loading Aspose.Cells | Aspose.Cells distinct error messages for password vs corruption

using System;
using System.IO;
using Aspose.Cells;

// Demonstrates loading an encrypted Excel file with Aspose.Cells in C#, checking file existence, applying LoadOptions with a password, and using CellsException message analysis to distinguish an incorrect password from a corrupted or unreadable workbook.
class Program
{
    static void Main()
    {
        // Path to the encrypted Excel file
        string filePath = "encrypted.xlsx";

        // Password supplied by the user
        string password = "userPassword";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file \"{filePath}\" was not found.");
            return;
        }

        // LoadOptions with the password for opening encrypted files
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = password
        };

        try
        {
            // Attempt to load the workbook using the provided password
            Workbook workbook = new Workbook(filePath, loadOptions);

            // If loading succeeds, continue processing the workbook here
            Console.WriteLine("Workbook loaded successfully.");
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws CellsException for both wrong password and corrupted files.
            // Distinguish based on the error message.

            // Wrong password scenario (message typically contains "Password")
            if (ex.Message.IndexOf("Password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("Error: Incorrect password provided.");
            }
            else
            {
                // Any other CellsException is treated as a corrupted or unreadable file
                Console.WriteLine("Error: The file appears to be corrupted or unreadable.");
            }
        }
        catch (Exception ex)
        {
            // Catch any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
        }
    }
}
