// Title: Check if an Excel workbook is password‑protected and log its encryption status using Aspose.Cells for .NET
// Common Searches: how to detect password protection on an .xlsx file with Aspose.Cells in C# | determine if an Excel workbook is encrypted without supplying a password using Aspose.Cells | log encryption status of a spreadsheet after loading it with Aspose.Cells | catch CellsException to identify protected Excel files in .NET | audit Excel file encryption state programmatically with Aspose.Cells
// Tags: Aspose.Cells detect encrypted workbook | C# check Excel file password protection | load Excel workbook without password Aspose.Cells | handle CellsException for protected spreadsheets | audit Excel encryption status .NET

using Aspose.Cells;
using System;
using System.IO;

// The example loads an Excel file with Aspose.Cells using LoadOptions without a password, catches a CellsException that indicates a required password, sets an encryption flag, and writes the file path and encryption status to the console for auditing purposes.
class Program
{
    static void Main()
    {
        // Path to the Excel file to be inspected
        string filePath = "sample.xlsx";

        // Verify that the file exists before attempting to load it
        if (!File.Exists(filePath))
        {
            Console.WriteLine($"Error: File not found - {filePath}");
            return;
        }

        // Prepare load options (no password supplied)
        LoadOptions loadOptions = new LoadOptions(LoadFormat.Xlsx);
        Workbook workbook = null;
        bool isEncrypted = false;

        try
        {
            // Attempt to load the workbook without a password
            workbook = new Workbook(filePath, loadOptions);
            // If loading succeeds, the file is not encrypted
            isEncrypted = false;
        }
        catch (CellsException ex)
        {
            // Aspose.Cells throws an exception when a password is required.
            // Detect encryption based on the exception message.
            if (!string.IsNullOrEmpty(ex.Message) &&
                ex.Message.IndexOf("password", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                isEncrypted = true;
            }
            else
            {
                // Other CellsException – rethrow or handle as needed
                Console.WriteLine($"CellsException: {ex.Message}");
                return;
            }
        }
        catch (Exception ex)
        {
            // Handle any other unexpected exceptions
            Console.WriteLine($"Unexpected error: {ex.Message}");
            return;
        }

        // Log the encryption status for auditing
        Console.WriteLine($"File: {filePath}");
        Console.WriteLine($"Encrypted: {isEncrypted}");
    }
}
