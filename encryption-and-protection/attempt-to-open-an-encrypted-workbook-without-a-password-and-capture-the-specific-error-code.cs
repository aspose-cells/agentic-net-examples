// Title: Detect encrypted Excel workbook by handling Aspose.Cells CellsException without providing a password in C#
// Common Searches: Aspose.Cells catch CellsException for encrypted workbook when password not supplied C# | How to determine if an Excel file is password protected using Aspose.Cells .NET | Retrieve error message from Aspose.Cells when opening encrypted .xlsx without a password | What exception does Aspose.Cells Workbook constructor throw for encrypted files | Check encryption status of an Excel file with Aspose.Cells without opening it
// Tags: Aspose.Cells catch CellsException encrypted workbook | Workbook constructor encrypted Excel .NET | detect password protected Excel Aspose.Cells | handle encryption error Aspose.Cells C# | CellsException error code encrypted file

using System;
using System.IO;
using Aspose.Cells;

// The sample verifies the existence of an encrypted.xlsx file, attempts to load it with the Workbook constructor without a password, catches the resulting CellsException, outputs the exception message, and checks the message for the term "encrypted" to confirm that the workbook requires a password.
class Program
{
    static void Main()
    {
        // Path to the encrypted workbook
        string filePath = "encrypted.xlsx";

        // Verify that the file exists to avoid FileNotFoundException
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"File not found: {filePath}");
            return;
        }

        try
        {
            // Attempt to load the workbook without providing a password
            Workbook workbook = new Workbook(filePath);
            Console.WriteLine("Workbook opened successfully (unexpected).");
        }
        catch (CellsException ex)
        {
            // Handle Aspose.Cells specific exceptions
            Console.WriteLine($"Failed to open workbook: {ex.Message}");

            // Simple check for encryption based on exception message
            if (ex.Message.IndexOf("encrypted", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                Console.WriteLine("The workbook is encrypted and requires a password.");
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"An unexpected error occurred: {ex.Message}");
        }
    }
}
