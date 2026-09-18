// Title: Check whether an Excel workbook is password‑protected with Aspose.Cells for .NET and log the result
// AI Prompts: Generate C# code that loads an .xlsx file using Aspose.Cells, catches the CellsException when the file is password‑protected, and prints "Encrypted" or "Not Encrypted" to the console. | Create a .NET console application that verifies a given Excel file exists, attempts to open it without a password via Aspose.Cells, determines its encryption status, and writes the status to a log output.
// Common Searches: asp.net how to detect password protection on an Excel file using Aspose.Cells | c# Aspose.Cells check if .xlsx is encrypted without providing a password | log encryption status of an Excel workbook with Aspose.Cells in a console app | catch CellsException to determine if an Excel file is password‑protected | determine Excel workbook encryption state using Aspose.Cells for .NET
// Tags: Aspose.Cells detect encrypted workbook | C# check Excel file password protection | log Excel encryption status .NET | handle CellsException for password‑protected file | audit Excel workbook encryption Aspose

using Aspose.Cells;
using System;
using System.IO;

// The example verifies the existence of a specified .xlsx file, attempts to load it with Aspose.Cells without a password, catches a CellsException containing the word "password" to identify encryption, and outputs the encryption status (Encrypted or Not Encrypted) to the console for auditing purposes.
class ExcelEncryptionChecker
{
    static void Main(string[] args)
    {
        // Path to the Excel file to be audited
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: The file '{filePath}' was not found.");
            return;
        }

        bool isEncrypted = false;

        try
        {
            // Attempt to load the workbook without providing a password
            Workbook workbook = new Workbook(filePath);
            // If loading succeeds, the file is not encrypted
            isEncrypted = false;
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws a CellsException when the file is password‑protected.
            // Check the exception message for password‑related text.
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isEncrypted = true;
            }
            else
            {
                // Re‑throw unexpected CellsException errors
                throw;
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
            return;
        }

        // Log the encryption status
        Console.WriteLine($"File '{filePath}' encryption status: {(isEncrypted ? "Encrypted" : "Not Encrypted")}");
    }
}
