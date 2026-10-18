// Title: Handle failed decryption of an encrypted Excel workbook with a wrong password using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells how to catch CellsException when opening a password‑protected XLSX with an invalid password | using LoadOptions to open encrypted Excel file with wrong password and retrieve error details in Aspose.Cells | Aspose.Cells C# example of handling decryption failure for a protected workbook | what exception is thrown by Aspose.Cells when workbook password is incorrect | debugging failed workbook load due to wrong password with Aspose.Cells C#
// Tags: Aspose.Cells LoadOptions incorrect password handling | CellsException capture decryption error C# | open encrypted XLSX with wrong password Aspose.Cells | error handling for password‑protected workbook Aspose.Cells | failed workbook decryption Aspose.Cells C#

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the encrypted Excel file exists, sets an intentionally wrong password in LoadOptions, attempts to load the workbook, and catches CellsException (and generic Exception) to display detailed error messages and stack traces, illustrating proper handling of decryption failures with Aspose.Cells in C#.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook file
        string filePath = "EncryptedWorkbook.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        // Intentionally incorrect password
        string wrongPassword = "wrongPassword";

        // Configure load options with the wrong password
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx)
        {
            Password = wrongPassword
        };

        try
        {
            // Attempt to open the encrypted workbook
            Workbook workbook = new Workbook(filePath, loadOptions);
            Console.WriteLine("Workbook opened successfully (this is unexpected with a wrong password).");
        }
        catch (CellsException ex)
        {
            // Capture and display Aspose.Cells specific exception details
            Console.WriteLine("Failed to open the encrypted workbook:");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }
        catch (Exception ex)
        {
            // Capture any other unexpected exceptions
            Console.WriteLine("An unexpected error occurred:");
            Console.WriteLine($"Message: {ex.Message}");
            Console.WriteLine($"Stack Trace: {ex.StackTrace}");
        }
    }
}
